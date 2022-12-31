import numpy as np
import pandas as pd
from string import capwords
import statistics as st
import dataframe_image as dfi

#Define a function to return the .csv data file name
def csv(report_type):
    file_names = {"mediahub": "MarkedStudent_DetailData",
                  "student app": "OneAppTeacherStudent",
                  "kpi": "ByTeacher_DetailData"}
    #Check for .csv files in folder
    possible_csvs = ["{}.csv".format(file_names[report_type])]
    for i in range(1,100):
        possible_csvs.append("{} ({}).csv".format(file_names[report_type], i))
    csvs = possible_csvs.copy()
    for i in possible_csvs:
        try:
            pd.read_csv(i, encoding='utf-16')
        except FileNotFoundError:
            csvs.remove(i)
    #Select a .csv file
    if len(csvs) == 1: file = csvs[0]
    elif len(csvs) == 0: file = 0
    else:
        print("Found {} .csv files:".format(len(csvs)))
        print(*csvs, sep='\n')
        while True:
            file = input("Please select one (by name or number) or type exit\n")
            if file in csvs: break
            elif int(file) in range(1,len(csvs)+1):
                file = csvs[int(file)-1]
                break
            elif file == "e" or file == "exit" or file == "Exit": break
            else:
                print("Invalid answer")
                continue
    return file

#Define function to return report name correctly
def reportname(df, report_type):
    if report_type == "mediahub": teacher_list = np.unique(df[["main_teacher", "co_teacher"]].values)
    if report_type == "student app":
        teacher_list = np.unique(df["Teacher"])
        teacher_list = np.ndarray.tolist(teacher_list)
    if report_type == "kpi": teacher_list = np.unique(df["Reservation Teacher"])
    while True:
        name = input("Please type 'summary', 'all', or enter a full name\n(type exit/e to exit)\n")
        if name.lower() in [x.lower() for x in teacher_list]: break
        elif name.lower() == "all" or name.lower() == "summary": break
        elif name.lower() == "e" or name.lower() == "exit": return 0
        else:
            print("Name not recognised. Please enter valid name.")
            continue
    return name.lower()

#Define function to generate MediaHub reports
def generate_mh_report(df, name):
    if name == 0:
        print("Exiting program")
        return 0
    df = df.sort_values(["main_teacher", "co_teacher", "group_code"])
    df.set_index("group_code", inplace=True)
    #Filter data frame for teacher's name
    if name != "all" and name != "summary":
        df = df[(df["main_teacher"].str.lower() == name) | (df["co_teacher"].str.lower() == name)]
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
            writer = pd.ExcelWriter("{} MH Report.xlsx".format(capwords(i)), engine='xlsxwriter')
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
        writer = pd.ExcelWriter("{} MH Report.xlsx".format(capwords(name)), engine='xlsxwriter')
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

#Define function to generate Student APP Usage reports
def generate_app_report(df, name):
    def overview(df, app_kpis):
        df_ov = df.copy()
        for kpi in app_kpis:
            df_ov[kpi] = df_ov[kpi].str.rstrip("%").astype(float)
        df_ov.drop(["Account Name", "Account English Name"], axis=1, inplace=True)
        ops = {"main_teacher": st.mode,
               "co_teacher": st.mode,
               "LG_Completion": np.nanmean,
               "HW_Completion": np.nanmean,
               "Read_Completion": np.nanmean,
               "Book_Read": np.nanmean,
               "Vocab_Completion": np.nanmean}
        df_ov = df_ov.groupby(["Group Code"], as_index=True).agg(ops).round(2)
        df_ov = df_ov.sort_values([ "main_teacher","co_teacher","Group Code"])
        return df_ov
    def kpis(df, app_kpis):
        df_list = [df.copy()]*len(app_kpis)
        hundreds = ["1", "100", "100%", "100.0%", "100.00%", "100.000%"]
        for i in range(len(app_kpis)):
            if app_kpis[i] == "LG_Completion":
                df_list[i] = df_list[i][df_list[i].index.str.contains("BSV")]
            for j in hundreds:
                df_list[i] = df_list[i][(df_list[i][app_kpis[i]] != j)]
            for k in app_kpis:
                if k != app_kpis[i]:
                    df_list[i] = df_list[i].drop(k, axis=1)
        return df_list

    if name == 0:
        print("Exiting program")
        return 0
    print("Generating Student APP Report")
    #Define Student APP KPIs
    app_kpis = ["LG_Completion", "HW_Completion",
                "Read_Completion", "Book_Read", "Vocab_Completion"]
    #Filter data frame for teacher's name
    df = df.sort_values(["Teacher", "main_teacher", "co_teacher", "Group Code"])
    if name != "all" and name != "summary":
        df = df[df["Teacher"].str.lower() == name]
        df.drop("Teacher", axis=1, inplace=True)
    df.set_index("Group Code", inplace=True)
    #df.drop_duplicates(keep=False,inplace=True)
    if name == "all":
        for teacher in np.unique(df["Teacher"]):
            df_teacher = df.copy()
            df_teacher = df_teacher[df_teacher["Teacher"] == teacher]
            df_teacher.drop("Teacher", axis=1, inplace=True)
            #Create Overview dataframe
            df_ov = overview(df_teacher, app_kpis)
            #Create list of KPI dataframes
            df_list = kpis(df, app_kpis)
            #Create sheets
            file_name = "{} Student APP Report.xlsx".format(capwords(teacher))
            format1 = {'type': '3_color_scale',
                          'min_value': '0',
                          'mid_value': '50',
                          'max_value': '100',
                          'min_color': '#FF0F0F',
                          'max_color': '#00F000'}
            writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            workbook = writer.book
            df_ov.to_excel(writer, sheet_name="Overview")
            worksheet_ov = writer.sheets["Overview"]
            for col in df_ov:
                    col_index = df_ov.columns.get_loc(col)
                    worksheet_ov.set_column(col_index, col_index, len(col))
            worksheet_ov.conditional_format('A1:I99', format1)

            worksheets = []
            for i in range(len(app_kpis)):
                df_list[i].to_excel(writer, sheet_name="{} Report".format(app_kpis[i]))
                worksheets.append(writer.sheets["{} Report".format(app_kpis[i])])
                for col in df_list[i]:
                        col_index = df_list[i].columns.get_loc(col)
                        worksheets[i].set_column(col_index, col_index, len(col))
                worksheets[i].conditional_format('A1:H999', {'type': 'text',
                                                 'criteria': 'not containing',
                                                 'value': '@',
                                                 'format': workbook.add_format({'bg_color': '#ff0000'})})
            writer.save()

    else:
        #Create Overview dataframe
        df_ov = overview(df, app_kpis)
        #Create list of KPI dataframes
        df_list = kpis(df, app_kpis)
        #Create sheets
        writer = pd.ExcelWriter("{} Student APP Report.xlsx".format(capwords(name)), engine='xlsxwriter')
        workbook = writer.book
        df_ov.to_excel(writer, sheet_name="Overview")
        worksheet_ov = writer.sheets["Overview"]
        for col in df_ov:
                col_index = df_ov.columns.get_loc(col)
                worksheet_ov.set_column(col_index, col_index, len(col))
        worksheet_ov.conditional_format('A1:I99', {'type': '3_color_scale',
                                                   'min_value': '0',
                                                   'mid_value': '50',
                                                   'max_value': '100',
                                                   'min_color': '#FF0F0F',
                                                   'mid_color': '#FFFF00',
                                                   'max_color': '#00F000'})

        worksheets = []
        for i in range(len(app_kpis)):
            df_list[i].to_excel(writer, sheet_name="{} Report".format(app_kpis[i]))
            worksheets.append(writer.sheets["{} Report".format(app_kpis[i])])
            for col in df_list[i]:
                    col_index = df_list[i].columns.get_loc(col)
                    worksheets[i].set_column(col_index, col_index, len(col))
            worksheets[i].conditional_format('A1:H999', {'type': 'text',
                                             'criteria': 'not containing',
                                             'value': '@',
                                             'format': workbook.add_format({'bg_color': '#ff0000'})})
        writer.save()

#Define function to generate KPI reports
def generate_kpi_report(df, name):
    if name == 0:
        print("Exiting program")
        return 0
    #GENERATE REPORT
    print("Generating KPI Report")
    #Replace NaN values with 0.0
    df.fillna(0.0, inplace=True)
    #Remove '%' from percentage values as '%' already specified in column names
    for i in list(filter(lambda x: "%" in x, list(df))): df[i] = df[i].str.rstrip("%").astype(float)
    #Filter out blank Group IDs
    df = df[df["Group ID"] != 0.0]
    #Set Group ID as Index
    df.set_index("Group ID", inplace=True)
    #Filter data by teacher name if applicable
    if name != "all" and name != "summary":
        df = df[df["Reservation Teacher"] == name]
        df.drop("Reservation Teacher", axis=1, inplace=True)
    #Generate KPI reports for 'all' teachers
    if name == "all":
        for teacher in np.unique(df["Reservation Teacher"]):
            #Create a copy of the dataframe, filtered for teacher
            df_tchr = df.copy()
            df_tchr = df_tchr[df_tchr["Reservation Teacher"] == teacher]
            df_tchr.drop("Reservation Teacher", axis=1, inplace=True)
            #Set Excel Writer
            writer = pd.ExcelWriter("{} KPI Report.xlsx".format(capwords(teacher)), engine='xlsxwriter')
            #Align cell text to center
            df_tchr.style.set_properties(**{'text-align': 'center'}).to_excel(writer, sheet_name="KPIs")
            #Set workbook and worksheet
            workbook, worksheet = writer.book, writer.sheets["KPIs"]
            workbook.add_format().set_text_wrap()
            #Set cell widths to fit text
            for column in df_tchr:
                column_index = df_tchr.columns.get_loc(column)
                worksheet.set_column(column_index, column_index, len(column))
            #Apply conditional formatting to the worksheets
            worksheet.conditional_format('B2:C500', {'type': '3_color_scale',
                                                    'min_value': '0',
                                                    'max_value': '124.3',
                                                    'min_color': '#FF0F0F',
                                                    'mid_color': '#FFFF00',
                                                    'max_color': '#00F000'})
            worksheet.conditional_format('D2:I500', {'type': '3_color_scale',
                                                    'min_value': '0',
                                                    'mid_value': '50',
                                                    'max_value': '100',
                                                    'min_color': '#FF0F0F',
                                                    'mid_color': '#FFFF00',
                                                    'max_color': '#00F000'})
            writer.save()
    else:
        #Set Excel Writer
        writer = pd.ExcelWriter("{} KPI Report.xlsx".format(capwords(name)), engine='xlsxwriter')
        #Align cell text to center
        #df.style.set_properties(**{'text-align': 'center'}).to_excel(writer, sheet_name="KPIs")
        df.to_excel(writer, sheet_name="KPIs")
        #Set workbook and worksheet
        workbook, worksheet = writer.book, writer.sheets["KPIs"]
        workbook.add_format().set_text_wrap()
        #Set cell widths to fit text
        for column in df:
            column_index = df.columns.get_loc(column)
            worksheet.set_column(column_index, column_index, len(column))
        #Apply conditional formatting to the worksheets
        worksheet.conditional_format('B2:C500', {'type': '3_color_scale',
                                                'min_value': '0',
                                                'max_value': '124.3',
                                                'min_color': '#FF0F0F',
                                                'mid_color': '#FFFF00',
                                                'max_color': '#00F000'})
        worksheet.conditional_format('D2:I500', {'type': '3_color_scale',
                                                'min_value': '0',
                                                'mid_value': '50',
                                                'max_value': '100',
                                                'min_color': '#FF0F0F',
                                                'mid_color': '#FFFF00',
                                                'max_color': '#00F000'})
        writer.save()
        if name != "summary":
            dfi.export(df, "df_image.png", table_conversion='matplotlib')
#Create a dictionary of the columns to include in each type of report
report_usecols = {"mediahub": ["group_code", "Student English Name", "main_teacher",
                               "co_teacher", "pa_name", "Marked Student",
                               "%Student ReceivedMedia File(Let'sTalk +Video)",
                               "%StudentReceived AcademicFeedback"],
                  "student app": ["Teacher", "Group Code", "Account Name",
                                  "Account English Name", "main_teacher",
                                  "co_teacher", "LG_Completion", "HW_Completion",
                                  "Read_Completion", "Book_Read", "Vocab_Completion"],
                  "kpi": ["Reservation Teacher", "Group ID",
                          "%Attendance MarkedOnTime", "TeachingTime (ACH)",
                          "%TeacherTaskCompletion", "%PT App Use",
                          "%SkillTestCompleted", "%Teacher-Led Skills Completion",
                          "%BS CanDo Completion"]}

#Select which type of report to generate
while True:
    print("Please choose a report type (or type 'exit'):")
    report_type = input("MediaHub\nStudent APP\nKPI\n").lower()
    if report_type in report_usecols:
        df = pd.read_csv(csv(report_type), sep='\t', usecols=report_usecols[report_type],
                         index_col=False, encoding='utf-16')
        break
    elif report_type == "e" or report_type == "exit": break
    else:
        print("Unrecognised report type.")
        continue

#Generate the selected report
if report_type == "mediahub": generate_mh_report(df, reportname(df, report_type))
if report_type == "student app": generate_app_report(df, reportname(df, report_type))
if report_type == "kpi": generate_kpi_report(df, reportname(df, report_type))
