import os
import excel2img
import pandas as pd
from docx.shared import Cm
from docxtpl import InlineImage


def create_report(context, doc, folderpath, file):
    file_Profile123 = file
    if os.path.exists(os.path.join(folderpath, file_Profile123)):
        file_Profile123 = os.path.join(folderpath, file_Profile123)
        # Generate Contests for Tab C
        excel2img.export_img(file_Profile123, 'KPIImages/Tables/' + 'C1_KPI_Results.png', "Output", "AB9:AG12")
        C1_KPI_Results = InlineImage(doc, 'KPIImages/Tables/C1_KPI_Results.png', width=Cm(17), height=Cm(2))

        excel2img.export_img(file_Profile123, 'KPIImages/Tables/' + 'C1_KPI_Table.png', "Output", "AB2:AI5")
        C1_KPI_Table = InlineImage(doc, 'KPIImages/Tables/C1_KPI_Table.png', width=Cm(17), height=Cm(3))

        # Generate Contests for Tab D1

        excel2img.export_img(file_Profile123, 'KPIImages/Tables/' + 'D1_Linksys_Table.png', "", "Output!B2:K9")
        excel2img.export_img(file_Profile123, 'KPIImages/Tables/' + 'D1_Asus_Table.png', "", "Output!B14:K21")
        excel2img.export_img(file_Profile123, 'KPIImages/Tables/' + 'D1_Nokia_Table.png', "", "Output!B26:K33")

        excel2img.export_img(file_Profile123, 'KPIImages/Graphs/' + 'D1_Linksys_Profile1_Graph.png',
                             "Linksys MOS Profile 1",
                             "S2:AB23")
        excel2img.export_img(file_Profile123, 'KPIImages/Graphs/' + 'D1_Linksys_Profile2_Graph.png',
                             "Linksys MOS Profile 2",
                             "S2:AB23")
        excel2img.export_img(file_Profile123, 'KPIImages/Graphs/' + 'D1_Linksys_Profile3_Graph.png',
                             "Linksys MOS Profile 3",
                             "S2:AB23")

        excel2img.export_img(file_Profile123, 'KPIImages/Graphs/' + 'D1_Asus_Profile1_Graph.png', "Asus MOS Profile 1",
                             "S2:AB23")
        excel2img.export_img(file_Profile123, 'KPIImages/Graphs/' + 'D1_Asus_Profile2_Graph.png', "Asus MOS Profile 2",
                             "S2:AB23")
        excel2img.export_img(file_Profile123, 'KPIImages/Graphs/' + 'D1_Asus_Profile3_Graph.png', "Asus MOS Profile 3",
                             "S2:AB23")

        excel2img.export_img(file_Profile123, 'KPIImages/Graphs/' + 'D1_Nokia_Profile1_Graph.png',
                             "Nokia MOS Profile 1",
                             "S2:AB23")
        excel2img.export_img(file_Profile123, 'KPIImages/Graphs/' + 'D1_Nokia_Profile2_Graph.png',
                             "Nokia MOS Profile 2",
                             "S2:AB23")
        excel2img.export_img(file_Profile123, 'KPIImages/Graphs/' + 'D1_Nokia_Profile3_Graph.png',
                             "Nokia MOS Profile 3",
                             "S2:AB23")

        D1_Linksys_Table = InlineImage(doc, 'KPIImages/Tables/D1_Linksys_Table.png', width=Cm(16), height=Cm(4))
        D1_Asus_Table = InlineImage(doc, 'KPIImages/Tables/D1_Asus_Table.png', width=Cm(16), height=Cm(4))
        D1_Nokia_Table = InlineImage(doc, 'KPIImages/Tables/D1_Nokia_Table.png', width=Cm(16), height=Cm(4))

        D1_Linksys_Profile1_Graph = InlineImage(doc, 'KPIImages/Graphs/D1_Linksys_Profile1_Graph.png', width=Cm(8),
                                                height=Cm(4))
        D1_Linksys_Profile2_Graph = InlineImage(doc, 'KPIImages/Graphs/D1_Linksys_Profile2_Graph.png', width=Cm(8),
                                                height=Cm(4))
        D1_Linksys_Profile3_Graph = InlineImage(doc, 'KPIImages/Graphs/D1_Linksys_Profile3_Graph.png', width=Cm(8),
                                                height=Cm(4))

        D1_Asus_Profile1_Graph = InlineImage(doc, 'KPIImages/Graphs/D1_Asus_Profile1_Graph.png', width=Cm(8),
                                             height=Cm(6))
        D1_Asus_Profile2_Graph = InlineImage(doc, 'KPIImages/Graphs/D1_Asus_Profile2_Graph.png', width=Cm(8),
                                             height=Cm(6))
        D1_Asus_Profile3_Graph = InlineImage(doc, 'KPIImages/Graphs/D1_Asus_Profile3_Graph.png', width=Cm(8),
                                             height=Cm(6))

        D1_Nokia_Profile1_Graph = InlineImage(doc, 'KPIImages/Graphs/D1_Nokia_Profile1_Graph.png', width=Cm(8),
                                              height=Cm(6))
        D1_Nokia_Profile2_Graph = InlineImage(doc, 'KPIImages/Graphs/D1_Nokia_Profile2_Graph.png', width=Cm(8),
                                              height=Cm(6))
        D1_Nokia_Profile3_Graph = InlineImage(doc, 'KPIImages/Graphs/D1_Nokia_Profile3_Graph.png', width=Cm(8),
                                              height=Cm(6))

        context.update({
            'D1_Linksys_Table': D1_Linksys_Table,
            'D1_Asus_Table': D1_Asus_Table,
            'D1_Nokia_Table': D1_Nokia_Table,
            'D1_Linksys_Profile1_Graph': D1_Linksys_Profile1_Graph,
            'D1_Linksys_Profile2_Graph': D1_Linksys_Profile2_Graph,
            'D1_Linksys_Profile3_Graph': D1_Linksys_Profile3_Graph,
            'D1_Asus_Profile1_Graph': D1_Asus_Profile1_Graph,
            'D1_Asus_Profile2_Graph': D1_Asus_Profile2_Graph,
            'D1_Asus_Profile3_Graph': D1_Asus_Profile3_Graph,
            'D1_Nokia_Profile1_Graph': D1_Nokia_Profile1_Graph,
            'D1_Nokia_Profile2_Graph': D1_Nokia_Profile2_Graph,
            'D1_Nokia_Profile3_Graph': D1_Nokia_Profile3_Graph,
            'C1_KPI_Results': C1_KPI_Results,
            'C1_KPI_Table': C1_KPI_Table
        })

        final_context = get_KPI_Color(context, doc, folderpath, file_Profile123)
        # print(final_context)
        # doc.render(final_context)
        # doc.save('../Reports/Profile123.docx')
        print("Profile123 C & D Saved!!")
        # time.sleep(3)
        # Cr.delete_files_in_folder('KPIImages/Graphs')
        # Cr.delete_files_in_folder('KPIImages/Tables')
    else:
        print('No File Handovers (Profile 1,2,3) Exists!')


def get_KPI_Color(context, doc, folderpath, file_Profile123):
    pd.set_option('display.max_columns', 20)
    df_Output = pd.read_excel(os.path.join(folderpath, file_Profile123), sheet_name='Output')
    AP = ['Linksys', 'Asus', 'Nokia']
    # print(str(df_Output.iat[4, 15]))
    for i, item in enumerate(AP):
        if i == 0:
            context = display_KPI_Results(context, doc, df_Output, item, 2, 7)
        elif i == 1:
            context = display_KPI_Results(context, doc, df_Output, item, 14, 19)
        elif i == 2:
            context = display_KPI_Results(context, doc, df_Output, item, 26, 31)
    return context


def display_KPI_Results(context, doc, df_Output, AP, start_i, end_i):
    # yellow = "00FFFF00"
    # red = '00FF0000'
    # green = '5cb800'
    Profile_Number = 1
    KPIs = ['Attempts', 'Call Drops', 'Handovers', 'RSRP', 'RSSI', 'MOS']
    for i in range(start_i, end_i, 2):
        # print('Profile_Number: ' + str(Profile_Number))
        KPI_Number = 0
        # print('i: ' + str(i))
        for j in range(14, 20):
            KPI_placeholder_name = 'C2_P' + str(Profile_Number) + '_row' + str(i) + '_col' + str(j)
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

# if __name__ == '__main__':
#     if os.path.exists('KPI Report.docx'):
#         os.remove('KPI Report.docx')
#     path = r'D:\Projects\WFC\Element Report'
#     # path = input("Enter the Report folder path: ")
#     device_name = input("Enter the Device Name: ")
#     context['Device_Name'] = device_name
#     create_report(context, path)
