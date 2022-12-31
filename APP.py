import numpy as np
import pandas as pd
import string as str
import statistics as st

def csv():
    #Verify .csv files in folder
    possible_csvs = ["OneAppTeacherStudent.csv"]
    for i in range(1,10):
        possible_csvs.append("OneAppTeacherStudent ({}).csv".format(i))
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
def username(df):
    while True:
        names_list = pd.unique(df["Teacher"])
        name = input("Please type 'summary', 'all', or enter a full name\n(type exit/e to exit)\n")
        if name.lower() in [x.lower() for x in names_list]: break
        elif name.lower() == "all" or name.lower() == "summary": break
        elif name.lower() == "e" or name.lower() == "exit": return 0
        else:
            print("Name not recognised. Please enter valid name.")
            continue
    return name.lower()
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
    #def create_excel(df=pd.DataFrame(), file_name="Workbook1.xlsx", sheet_name='Sheet1', format=False,
    #                 format_range='A1:Z99', set_col_lengths=True, engine='xlsxwriter'):
    #    writer = pd.ExcelWriter(file_name, engine=engine)
    #    workbook = writer.book
    #    df.to_excel(writer, sheet_name=sheet_name)
    #    worksheet = writer.sheets[sheet_name]
    #    if set_col_lengths == True:
    #        for col in df:
    #            col_index = df.columns.get_loc(col)
    #            worksheet.set_column(col_index, col_index, len(col))
    #    if format != False:
    #        worksheet.conditional_format(format_range, format)
    #    return writer, workbook, worksheet

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
        df = df[df["Teacher"] == name]
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
            file_name = "{} Student APP Report.xlsx".format(str.capwords(teacher))
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

            #create_excel(df_ov, file_name=file_name, sheet_name="Overview", format=format1)

            #for i in range(len(app_kpis)):
            #    writer, workbook, worksheet = create_excel(df_list[i], file_name=file_name,
            #                sheet_name="{} Report".format(app_kpis[i]))
            #    format2 = {'type': 'text',
            #               'criteria': 'not containing',
            #               'value': '@',
            #               'format': workbook.add_format({'bg_color': '#ff0000'})}
            #    worksheet.conditional_format('A1:H999', format2)

    else:
        #Create Overview dataframe
        df_ov = overview(df, app_kpis)
        #Create list of KPI dataframes
        df_list = kpis(df, app_kpis)
        #Create sheets
        writer = pd.ExcelWriter("{} Student APP Report.xlsx".format(str.capwords(name)), engine='xlsxwriter')
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

try:
    #Choose which columns to use
    usecols = ["Teacher", "Group Code", "Account Name", "Account English Name",
               "main_teacher", "co_teacher", "LG_Completion",
               "HW_Completion", "Read_Completion", "Book_Read", "Vocab_Completion"]
    #Read data into pandas data frames
    df = pd.read_csv(csv(), sep='\t', usecols=usecols,
                     index_col=False, encoding='utf-16')
    df["Teacher"] = df["Teacher"].str.lower()
    generate_app_report(df, username(df))
except (FileNotFoundError, ValueError):
    print("No .csv files found. Exiting program")
#except NameError:
#    print("Exiting program")
