import os
from datetime import datetime
from docxtpl import DocxTemplate
from word_automation.ProfileScripts import Profile911 as P911, Profile123 as P123, Baseline as BL, \
    Profile4_IP_Impairment as P4, Profile56 as P56

context = {}
doc = DocxTemplate('KPI Template.docx')


def delete_files_in_folder(folder_path):
    """Deletes all files in the specified folder."""

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted: {file_path}")
        except Exception as e:
            print(f"Error deleting {file_path}: {e}")


if __name__ == '__main__':
    path = r'D:\Projects\WFC\SamsungA16\Element Report'
    # path = input("Enter the Report folder path: ")
    file_names = []
    if os.path.exists(path):
        if os.path.exists('Final WFC Report.docx'):
            os.remove('Final WFC Report.docx')
            file_names = [files for files in os.listdir(path) if files.endswith(".xlsx")]
        # print(file_names)
        delete_files_in_folder('KPIImages/Graphs')
        delete_files_in_folder('KPIImages/Tables')
        # device_name = input("Enter the Device Name: ")
        # software_version = input("Enter the Software Version: ")
        # IMEI = input("Enter the Device IMEI: ")

        device_name = 'Samsung A16'
        software_version = 'VVK35.54'
        IMEI = '352291420059098'

        context.update({
            'device_name': device_name,
            'tester': 'Raj Dinesh Kapadia',
            'report_date': datetime.now().strftime("%m/%d/%Y"),
            'version': '1',
            'software_version': str(software_version),
            'IMEI': str(IMEI)
        })

        for file in file_names:
            if str(file).__contains__('Baseline'):
                BL.get_baselline_image(context, doc, path, file)
            if str(file).__contains__('9.1.1'):
                P911.create_report(context, doc, path, file)
            elif str(file).__contains__('Handovers '):
                P123.create_report(context, doc, path, file)
            elif str(file).__contains__('Impairment'):
                P4.create_report(context, doc, path, file)
            elif str(file).__contains__('Rove'):
                P56.create_report(context, doc, path, file)
        # print(context)
        # file_paths = []
        # for root, _, files in os.walk('../Reports'):
        #     for file in files:
        #         file_paths.append(os.path.join(root, file))
        # doc = combine_word_documents(file_paths)
        doc.render(context)
        doc.save("Final WFC Report.docx")
        # delete_files_in_folder('KPIImages/Graphs')
        # delete_files_in_folder('KPIImages/Tables')
    else:
        print('No such Folder Exists!')
