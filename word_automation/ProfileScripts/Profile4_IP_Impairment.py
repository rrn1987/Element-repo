import os
import excel2img
import pandas as pd
from docx.shared import Cm
from docxtpl import InlineImage


def create_report(context, doc, folderpath, file):
    file_Profile4 = file
    if os.path.exists(os.path.join(folderpath, file_Profile4)):
        # Generate Contests for Tab E1
        file_Profile4 = os.path.join(folderpath, file_Profile4)
        excel2img.export_img(file_Profile4, 'KPIImages/Tables/' + 'E1_KPI_Results.png', "Output", "AJ2:AO4")
        E1_KPI_Results = InlineImage(doc, 'KPIImages/Tables/E1_KPI_Results.png', width=Cm(17), height=Cm(4))

        excel2img.export_img(file_Profile4, 'KPIImages/Tables/' + 'E1_KPI_Table.png', "Output", "AA2:AH5")
        E1_KPI_Table = InlineImage(doc, 'KPIImages/Tables/E1_KPI_Table.png', width=Cm(17), height=Cm(3))

        context.update({
            'E1_KPI_Results': E1_KPI_Results,
            'E1_KPI_Table': E1_KPI_Table
        })

        final_context = get_KPI_Color(context, doc, folderpath, file_Profile4)
        # print(context)
        # doc.render(context)
        # doc.save('../Reports/Profile911.docx')
        print("Impairment (Profile 4) E1 & E2 Report Saved!!")
        # Report for Tab F1
        create_report_IPImpairment(context, doc, folderpath, file_Profile4)
        print("Profile4 F Saved!!")
        # time.sleep(3)
        # Cr.delete_files_in_folder('KPIImages/Tables')
    else:
        print('No Impairment (Profile 4) Exists!')


def get_KPI_Color(context, doc, folderpath, file_Profile4):
    pd.set_option('display.max_columns', 20)
    df_Output = pd.read_excel(os.path.join(folderpath, file_Profile4), sheet_name='Output')
    # print(df_Output['Call Setup Time'].dropna())
    AP = ['Linksys', 'Asus', 'Google_Nest']
    # print(str(df_Output.iat[2, 23]))
    for i, item in enumerate(AP):
        if i == 0:
            context = display_KPI_Results(context, doc, df_Output, item, 2, 3)
        elif i == 1:
            context = display_KPI_Results(context, doc, df_Output, item, 10, 11)
        elif i == 2:
            context = display_KPI_Results(context, doc, df_Output, item, 18, 19)
    return context


def display_KPI_Results(context, doc, df_Output, AP, start_i, end_i):
    Profile_Number = 4
    KPIs = ['Attempts', 'Call Drops', 'Handovers', 'MOS before handover', 'MOS during and after handover',
            'Handover Delay impact to speech', 'Packet Loss']
    for i in range(start_i, end_i, 2):
        # print('Profile_Number: ' + str(Profile_Number))
        KPI_Number = 0
        # print('i: ' + str(i))
        for j in range(17, 24):
            KPI_placeholder_name = 'E2_P' + str(Profile_Number) + '_row' + str(i) + '_col' + str(j)
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
            elif result == 'nan':
                grey = InlineImage(doc, 'KPIImages/KPIColorCode/Grey_NA.JPG', width=Cm(1.5))
                context[KPI_placeholder_name] = grey
            KPI_Number += 1
    return context


def create_report_IPImpairment(context, doc, folderpath, file_Profile4):
    # Generate Contests for Tab F1

    excel2img.export_img(file_Profile4, 'KPIImages/Tables/' + 'F1_Linksys_Table.png', "", "Output!B2:N5")
    excel2img.export_img(file_Profile4, 'KPIImages/Tables/' + 'F1_Asus_Table.png', "", "Output!B10:N13")
    excel2img.export_img(file_Profile4, 'KPIImages/Tables/' + 'F1_Google_Table.png', "", "Output!B18:N21")

    excel2img.export_img(file_Profile4, 'KPIImages/Graphs/' + 'F1_Linksys_Profile4_Graph.png',
                         "Real Linksys MOS Profile 4",
                         "S2:AB23")

    excel2img.export_img(file_Profile4, 'KPIImages/Graphs/' + 'F1_Asus_Profile4_Graph.png', "Asus MOS Profile 4",
                         "S2:AB23")

    excel2img.export_img(file_Profile4, 'KPIImages/Graphs/' + 'F1_Google_Profile4_Graph.png', "Google MOS Profile 4",
                         "S2:AB23")

    F1_Linksys_Table = InlineImage(doc, 'KPIImages/Tables/F1_Linksys_Table.png', width=Cm(16), height=Cm(2))
    F1_Asus_Table = InlineImage(doc, 'KPIImages/Tables/F1_Asus_Table.png', width=Cm(16), height=Cm(2))
    F1_Google_Table = InlineImage(doc, 'KPIImages/Tables/F1_Google_Table.png', width=Cm(16), height=Cm(2))

    F1_Linksys_Profile4_Graph = InlineImage(doc, 'KPIImages/Graphs/F1_Linksys_Profile4_Graph.png', width=Cm(10),
                                            height=Cm(8))
    F1_Asus_Profile4_Graph = InlineImage(doc, 'KPIImages/Graphs/F1_Asus_Profile4_Graph.png', width=Cm(10),
                                            height=Cm(8))
    F1_Google_Profile4_Graph = InlineImage(doc, 'KPIImages/Graphs/F1_Google_Profile4_Graph.png', width=Cm(10),
                                            height=Cm(8))

    context.update({
        'F1_Linksys_Table': F1_Linksys_Table,
        'F1_Asus_Table': F1_Asus_Table,
        'F1_Google_Table': F1_Google_Table,
        'F1_Linksys_Profile4_Graph': F1_Linksys_Profile4_Graph,
        'F1_Asus_Profile4_Graph': F1_Asus_Profile4_Graph,
        'F1_Google_Profile4_Graph': F1_Google_Profile4_Graph,
    })
