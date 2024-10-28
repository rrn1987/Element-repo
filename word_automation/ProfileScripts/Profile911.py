import os
import excel2img
import pandas as pd
from docx.shared import Cm
from docxtpl import InlineImage
import win32com.client as win32


def create_report(context, doc, folderpath, file):
    file_Profile911 = file
    if os.path.exists(os.path.join(folderpath, file_Profile911)):
        # Generate Contests for Tab A1
        file_Profile911 = os.path.join(folderpath, file_Profile911)

        excel2img.export_img(file_Profile911, 'KPIImages/Tables/' + 'A1_KPI_Results.png', "Output", "AD2:AF4")
        A1_KPI_Results = InlineImage(doc, 'KPIImages/Tables/A1_KPI_Results.png', width=Cm(5), height=Cm(2))

        excel2img.export_img(file_Profile911, 'KPIImages/Tables/' + 'A1_AVG_KPI.png', "Output", "AH2:AJ5")
        A1_AVG_KPI = InlineImage(doc, 'KPIImages/Tables/A1_AVG_KPI.png', width=Cm(5), height=Cm(2))

        excel2img.export_img(file_Profile911, 'KPIImages/Tables/' + 'A1_KPI_Table.png', "Output", "S2:AA9")
        A1_KPI_Table = InlineImage(doc, 'KPIImages/Tables/A1_KPI_Table.png', width=Cm(17), height=Cm(3))

        excel2img.export_img(file_Profile911, 'KPIImages/Tables/' + 'B1_Linksys_Table.png', "", "Output!B2:L9")
        excel2img.export_img(file_Profile911, 'KPIImages/Tables/' + 'B1_Asus_Table.png', "", "Output!B14:L21")
        excel2img.export_img(file_Profile911, 'KPIImages/Tables/' + 'B1_Nest_Table.png', "", "Output!B26:L33")

        # get_charts(file_Profile911)

        # excel2img.export_img(file_Profile911, 'KPIImages/Graphs/' + 'B1_Linksys_Initiation_Graph.png',
        #                      "Output",
        #                      "Chart 1")
        # excel2img.export_img(file_Profile911, 'KPIImages/Graphs/' + 'D1_Linksys_Profile2_Graph.png',
        #                      "Linksys MOS Profile 2",
        #                      "S2:AB23")
        #
        # excel2img.export_img(file_Profile911, 'KPIImages/Graphs/' + 'D1_Asus_Profile1_Graph.png', "Asus MOS Profile 1",
        #                      "S2:AB23")
        # excel2img.export_img(file_Profile911, 'KPIImages/Graphs/' + 'D1_Asus_Profile2_Graph.png', "Asus MOS Profile 2",
        #                      "S2:AB23")
        #
        # excel2img.export_img(file_Profile911, 'KPIImages/Graphs/' + 'D1_Nokia_Profile1_Graph.png',
        #                      "Nokia MOS Profile 1",
        #                      "S2:AB23")
        # excel2img.export_img(file_Profile911, 'KPIImages/Graphs/' + 'D1_Nokia_Profile2_Graph.png',
        #                      "Nokia MOS Profile 2",
        #                      "S2:AB23")

        B1_Linksys_Table = InlineImage(doc, 'KPIImages/Tables/B1_Linksys_Table.png', width=Cm(16), height=Cm(4))
        B1_Asus_Table = InlineImage(doc, 'KPIImages/Tables/B1_Asus_Table.png', width=Cm(16), height=Cm(4))
        B1_Nest_Table = InlineImage(doc, 'KPIImages/Tables/B1_Nest_Table.png', width=Cm(16), height=Cm(4))

        # B1_Linksys_Initiation_Graph = InlineImage(doc, 'KPIImages/Graphs/B1_Linksys_Initiation_Graph.png', width=Cm(8),
        #                                           height=Cm(4))
        # B1_Linksys_Retention_Graph = InlineImage(doc, 'KPIImages/Graphs/B1_Linksys_Retention_Graph.png', width=Cm(8),
        #                                          height=Cm(4))
        #
        # B1_Asus_Initiation_Graph = InlineImage(doc, 'KPIImages/Graphs/B1_Asus_Initiation_Graph.png', width=Cm(8),
        #                                        height=Cm(6))
        # B1_Asus_Retention_Graph = InlineImage(doc, 'KPIImages/Graphs/B1_Asus_Retention_Graph.png', width=Cm(8),
        #                                       height=Cm(6))
        #
        # B1_Nest_Initiation_Graph = InlineImage(doc, 'KPIImages/Graphs/B1_Nest_Initiation_Graph.png', width=Cm(8),
        #                                        height=Cm(6))
        # B1_Nest_Retention_Graph = InlineImage(doc, 'KPIImages/Graphs/B1_Nest_Retention_Graph.png', width=Cm(8),
        #                                       height=Cm(6))

        context.update({
            'A1_KPI_Results': A1_KPI_Results,
            'A1_AVG_KPI': A1_AVG_KPI,
            'A1_KPI_Table': A1_KPI_Table,
            'B1_Linksys_Table': B1_Linksys_Table,
            'B1_Asus_Table': B1_Asus_Table,
            'B1_Nest_Table': B1_Nest_Table
            # 'B1_Linksys_Initiation_Graph': B1_Linksys_Initiation_Graph,
            # 'B1_Linksys_Retention_Graph': B1_Linksys_Retention_Graph,
            # 'B1_Asus_Initiation_Graph': B1_Asus_Initiation_Graph,
            # 'B1_Asus_Retention_Graph': B1_Asus_Retention_Graph,
            # 'B1_Nest_Initiation_Graph': B1_Nest_Initiation_Graph,
            # 'B1_Nest_Retention_Graph': B1_Nest_Retention_Graph,
        })

        final_context = get_KPI_Color(context, doc, folderpath, file_Profile911)
        # print(context)
        # doc.render(context)
        # doc.save('../Reports/Profile911.docx')
        print("Profile9.1.1 A1 & A2 Report Saved!!")
        print("Profile4 F Saved!!")
        # time.sleep(3)
        # Cr.delete_files_in_folder('KPIImages/Tables')
    else:
        print('No File 9.1.1 - offical Exists!')


def get_charts(file_Profile911):
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(file_Profile911)
    sheet = wb.Sheets('Output')
    chart_object = sheet.ChartObjects(1)
    chart = chart_object.Chart
    chart.Export(r'path_to_save_image\chart_image.png')
    wb.Close(SaveChanges=False)
    excel.Quit()


def get_KPI_Color(context, doc, folderpath, file_Profile911):
    pd.set_option('display.max_columns', 20)
    df_Output = pd.read_excel(os.path.join(folderpath, file_Profile911), sheet_name='Output')
    # print(df_Output['Call Setup Time'].dropna())
    AP = ['Linksys', 'Asus', 'Google_Nest']
    # print(str(df_Output.iat[2, 15]))
    for i, item in enumerate(AP):
        if i == 0:
            context = display_KPI_Results(context, doc, df_Output, item, 2, 7)
        elif i == 1:
            context = display_KPI_Results(context, doc, df_Output, item, 14, 19)
        elif i == 2:
            context = display_KPI_Results(context, doc, df_Output, item, 26, 31)
    return context


def display_KPI_Results(context, doc, df_Output, AP, start_i, end_i):
    Profile_Number = 1
    KPIs = ['Call Initiation', 'Call Retention', 'Call Setup Time', 'MOS']
    for i in range(start_i, end_i, 2):
        # print('Profile_Number: ' + str(Profile_Number))
        KPI_Number = 0
        # print('i: ' + str(i))
        for j in range(13, 17):
            KPI_placeholder_name = 'A2_P' + str(Profile_Number) + '_row' + str(i) + '_col' + str(j)
            # print(KPI_placeholder_name)
            result = str(df_Output.iat[i, j])
            # print(AP + " Profile" + str(Profile_Number) + ' ' + KPIs[KPI_Number] + ' ' + result)
            if result == 'Pass':
                green = InlineImage(doc, 'KPIImages/KPIColorCode/Green.JPG', width=Cm(1.5))
                context[KPI_placeholder_name] = green
            elif result == 'Fail':
                red = InlineImage(doc, 'KPIImages/KPIColorCode/Red.JPG', width=Cm(1.5))
                context[KPI_placeholder_name] = red
            elif result == 'Marginal':
                yellow = InlineImage(doc, 'KPIImages/KPIColorCode/Yellow.JPG', width=Cm(1.5))
                context[KPI_placeholder_name] = yellow
            elif result == 'Outperformed':
                magenta = InlineImage(doc, 'KPIImages/KPIColorCode/Magenta.JPG', width=Cm(1.5))
                context[KPI_placeholder_name] = magenta
            KPI_Number += 1
        Profile_Number += 1
    return context
