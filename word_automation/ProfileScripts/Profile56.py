import os
import excel2img
import pandas as pd
from docx.shared import Cm
from docxtpl import InlineImage


def create_report(context, doc, folderpath, file):
    file_Profile56 = file
    if os.path.exists(os.path.join(folderpath, file_Profile56)):
        file_Profile56 = os.path.join(folderpath, file_Profile56)
        # Generate Contests for Tab G
        excel2img.export_img(file_Profile56, 'KPIImages/Tables/' + 'G1_KPI_Results.png', "Output", "AH2:AM4")
        G1_KPI_Results = InlineImage(doc, 'KPIImages/Tables/G1_KPI_Results.png', width=Cm(17), height=Cm(3))

        excel2img.export_img(file_Profile56, 'KPIImages/Tables/' + 'G1_KPI_Table.png', "Output", "Y2:AF5")
        G1_KPI_Table = InlineImage(doc, 'KPIImages/Tables/G1_KPI_Table.png', width=Cm(17), height=Cm(3))

        # Generate Contests for Tab H1

        excel2img.export_img(file_Profile56, 'KPIImages/Tables/' + 'H1_Linksys_Table.png', "", "Output!B2:K7")
        excel2img.export_img(file_Profile56, 'KPIImages/Tables/' + 'H1_Asus_Table.png', "", "Output!B12:K17")
        excel2img.export_img(file_Profile56, 'KPIImages/Tables/' + 'H1_Nokia_Table.png', "", "Output!B22:K27")

        excel2img.export_img(file_Profile56, 'KPIImages/Graphs/' + 'H1_Linksys_Profile5_Graph.png',
                             "Linksys MOS Profile 5",
                             "S2:AB23")
        excel2img.export_img(file_Profile56, 'KPIImages/Graphs/' + 'H1_Linksys_Profile6_Graph.png',
                             "Linksys MOS Profile 6",
                             "S2:AB23")

        excel2img.export_img(file_Profile56, 'KPIImages/Graphs/' + 'H1_Asus_Profile5_Graph.png', "Asus MOS Profile 5",
                             "S2:AB23")
        excel2img.export_img(file_Profile56, 'KPIImages/Graphs/' + 'H1_Asus_Profile6_Graph.png', "Asus MOS Profile 6",
                             "S2:AB23")

        excel2img.export_img(file_Profile56, 'KPIImages/Graphs/' + 'H1_Nokia_Profile5_Graph.png', "Nokia MOS Profile 5",
                             "S2:AB23")
        excel2img.export_img(file_Profile56, 'KPIImages/Graphs/' + 'H1_Nokia_Profile6_Graph.png', "Nokia MOS Profile 6",
                             "S2:AB23")

        H1_Linksys_Table = InlineImage(doc, 'KPIImages/Tables/H1_Linksys_Table.png', width=Cm(16), height=Cm(4))
        H1_Asus_Table = InlineImage(doc, 'KPIImages/Tables/H1_Asus_Table.png', width=Cm(16), height=Cm(4))
        H1_Nokia_Table = InlineImage(doc, 'KPIImages/Tables/H1_Nokia_Table.png', width=Cm(16), height=Cm(4))

        H1_Linksys_Profile1_Graph = InlineImage(doc, 'KPIImages/Graphs/H1_Linksys_Profile5_Graph.png', width=Cm(8),
                                                height=Cm(6))
        H1_Linksys_Profile2_Graph = InlineImage(doc, 'KPIImages/Graphs/H1_Linksys_Profile6_Graph.png', width=Cm(8),
                                                height=Cm(6))

        H1_Asus_Profile1_Graph = InlineImage(doc, 'KPIImages/Graphs/H1_Asus_Profile5_Graph.png', width=Cm(8),
                                             height=Cm(6))
        H1_Asus_Profile2_Graph = InlineImage(doc, 'KPIImages/Graphs/H1_Asus_Profile6_Graph.png', width=Cm(8),
                                             height=Cm(6))

        H1_Nokia_Profile1_Graph = InlineImage(doc, 'KPIImages/Graphs/H1_Nokia_Profile5_Graph.png', width=Cm(8),
                                              height=Cm(6))
        H1_Nokia_Profile2_Graph = InlineImage(doc, 'KPIImages/Graphs/H1_Nokia_Profile6_Graph.png', width=Cm(8),
                                              height=Cm(6))

        context.update({
            'H1_Linksys_Table': H1_Linksys_Table,
            'H1_Asus_Table': H1_Asus_Table,
            'H1_Nokia_Table': H1_Nokia_Table,
            'H1_Linksys_Profile5_Graph': H1_Linksys_Profile1_Graph,
            'H1_Linksys_Profile6_Graph': H1_Linksys_Profile2_Graph,
            'H1_Asus_Profile5_Graph': H1_Asus_Profile1_Graph,
            'H1_Asus_Profile6_Graph': H1_Asus_Profile2_Graph,
            'H1_Nokia_Profile5_Graph': H1_Nokia_Profile1_Graph,
            'H1_Nokia_Profile6_Graph': H1_Nokia_Profile2_Graph,
            'G1_KPI_Results': G1_KPI_Results,
            'G1_KPI_Table': G1_KPI_Table
        })

        final_context = get_KPI_Color(context, doc, folderpath, file_Profile56)
        # print(final_context)
        # doc.render(final_context)
        # doc.save('../Reports/Profile123.docx')
        print("Profile56 G & H Saved!!")
        # time.sleep(3)
        # Cr.delete_files_in_folder('KPIImages/Graphs')
        # Cr.delete_files_in_folder('KPIImages/Tables')
    else:
        print('No File Handovers (Profile 1,2,3) Exists!')


def get_KPI_Color(context, doc, folderpath, file_Profile56):
    pd.set_option('display.max_columns', 20)
    df_Output = pd.read_excel(os.path.join(folderpath, file_Profile56), sheet_name='Output')
    AP = ['Linksys', 'Asus', 'Nokia']
    # print(str(df_Output.iat[2, 15]))
    for i, item in enumerate(AP):
        if i == 0:
            context = display_KPI_Results(context, doc, df_Output, item, 2, 5)
        elif i == 1:
            context = display_KPI_Results(context, doc, df_Output, item, 12, 15)
        elif i == 2:
            context = display_KPI_Results(context, doc, df_Output, item, 22, 25)
    return context


def display_KPI_Results(context, doc, df_Output, AP, start_i, end_i):
    Profile_Number = 5
    KPIs = ['Attempts', 'Call Drops', 'RSSI', 'RSRP', 'MOS before handover', 'MOS during and after handover',
            'Handovers occured', 'Handover Impact']
    for i in range(start_i, end_i, 2):
        # print('Profile_Number: ' + str(Profile_Number))
        KPI_Number = 0
        # print('i: ' + str(i))
        for j in range(14, 22):
            KPI_placeholder_name = 'G2_P' + str(Profile_Number) + '_row' + str(i) + '_col' + str(j)
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
