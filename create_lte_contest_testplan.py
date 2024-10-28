import os
import pandas as pd
import xml.etree.ElementTree as ET

Protocol_IMS = 'Protocol_IMS'
Protocol_VoLTE = 'Protocol_VoLTE'
Protocol_IR94 = 'Protocol_IR94'
Subject = [Protocol_IMS, Protocol_VoLTE, Protocol_IR94]
Device_Validation_Test_Platform_Complete = 'Device Validation_Test Platform'
Device_Validation_Test_Platform = 'DEVICE VALIDATION TEST PLATFORM'
Platform_Complete = 'Rohde & Schwarz CMX500'
Platform = 'Rohde & Schwarz'
Device_Validation_Category = 'Protocol_5G SA'
RESULT_Remaining = 'NO RUN'
VERSION = '23.40.19.11860'

# ATE_section = ['Ims-KAF53-23.40.19']
ATE_section = ['Ims-KAF50-23.40.19', 'Ims-KAF52-23.40.19', 'Ims-KAF53-23.40.19', 'Ims-KAF54-23.40.19',
               'Ims-KAF56-23.40.19', 'Ims-KAF57-23.40.19', 'Ims-KAO52-23.40.19', 'Ims-KAO53-23.40.19',
               'Ims-KAO54-23.40.19']

# Protocol_IMS & Protocol_VoLTE
#               'Ims-KAF50-23.40.19' - VoLTE,
#               'Ims-KAF52-23.40.19' - VoLTE EVS
#               'Ims-KAF53-23.40.19' - Error Codes
#               'Ims-KAF54-23.40.19' - VoLTE Phase 2
#               'Ims-KAO52-23.40.19' - VoLTE Phase 3
#               'Ims-KAO53-23.40.19' - VoLTE Phase 4
# Protocol_IR94 - 'Ims-KAF56-23.40.19', 'Ims-KAF57-23.40.19' , 'Ims-KAO54-23.40.19'
master_path = r'C:\ProgramData\Rohde-Schwarz\Contest\19\Testplans\Ims\ATE_TMO'


def aqcc_get_project_TC_Id(path, file):
    count = 0
    testcase_list = []
    df = pd.read_excel(os.path.join(path, file), sheet_name='Detail')
    # df = df.loc[df[Device_Validation_Test_Platform].str.contains(Platform, na=False)]
    df = df.loc[df['SUBJECT'].str.startswith((tuple(Subject)), na=False)]
    for id, row in df.iterrows():
        if row['RESULT'] == 'NO RUN':
            # print(row['ID'] + ' - ' + row['NAME'] + ' - ' + row['RESULT'])
            count += 1
            testcase_list.append(row['ID'])
    print('Total TCs required to be selected in the Project for Automation - ' + str(len(testcase_list)))
    # print(str(Id['ID']) + ' - ' + df['NAME'])
    tc_id_list, tc_id_not_in_list = map_Id_TestCase(testcase_list)
    for ids, row in df.iterrows():
        if row['ID'] in tc_id_not_in_list:
            print(row['ID'] + ' - ' + row['NAME'] + ' - Not Found in R&S Sheet')
    map_AllTestCases_xml(tc_id_list)


def map_Id_TestCase(testcase_list):
    tc_id_list = []
    tc_id_not_in_list = []
    count = 0
    df = pd.read_excel(os.path.join(os.getcwd(), 'TMO_NPT_ATE_PQA_Q3_2023_Release.xlsx'), sheet_name='Data')
    tid_tcname_dict = df.set_index('Test ID')['Test Case'].to_dict()
    for Id in testcase_list:
        if Id in tid_tcname_dict:
            count += 1
            # print(str(count) + ' - ' + Id + ' - ' + str(tid_tcname_dict[Id]))
            tc_id_list.append(str(tid_tcname_dict[Id]))
        else:
            tc_id_not_in_list.append(Id)
    # print('Total TCs Selected in ATE Sheet - ' + str(len(tc_id_list)))
    # print(tc_id_list)
    return tc_id_list, tc_id_not_in_list


def map_AllTestCases_xml(tc_id_list):
    rstt_name_list = []
    file_Alltestcases = os.path.join(os.getcwd(), "Alltestcases.xml")
    tree = ET.parse(file_Alltestcases)
    root = tree.getroot()
    for item in root[0].findall('Testcase'):
        rstt_name = str(item.get('Name'))
        testcase = item.find('Name').text
        if testcase in tc_id_list:
            rstt_name_list.append(rstt_name)
            tc_id_list.remove(testcase)
    print('Total TCs Selected in XML File - ' + str(len(rstt_name_list)))
    # print(*rstt_name_list)
    # if tc_id_list:
    #     print(str(*tc_id_list) + ' - Not found in XML File')
    create_testplan(rstt_name_list)


def create_testplan(rstt_name_list):
    # project_name = input("Enter the Project name: ")
    # shutil.copy(master_file, test_plan)
    each_section = ''
    count = 0
    total = 0
    for each_section in ATE_section:
        test_plan = master_path + '\\' + each_section + '.rstt'
        disable_count = disable_all_test_cases(test_plan)
        # print('Total TCs Disabled for ' + each_section + ' - ' + str(disable_count))
        count = enable_test_cases(test_plan, rstt_name_list)
        print('Total TCs enabled for ' + each_section + ' - ' + str(count))
        total += count
    print('Total TCs Enabled:' + str(total))
    if rstt_name_list:
        print('TCs Not Selected: ' + str(len(rstt_name_list)))
        print(*rstt_name_list)


def enable_test_cases(test_plan, rstt_name_list):
    tree = ET.parse(test_plan)
    root = tree.getroot()
    count = 0
    step_path = './Testplan/Sequence'
    test_path = step_path + '/Step'
    for item in root.findall(test_path):
        testcase = str(item.find('Test').get('Name'))
        version = str(item.find('Test').get('Version'))
        # print(testcase)   # + '- ' + disabled
        if testcase in rstt_name_list and version == VERSION:
            count += 1
            rstt_name_list.remove(testcase)
            # print(str(count) + '. ' + testcase + ' - Enabled')
            item.set('Disabled', 'False')
    tree.write(test_plan)
    return count


def disable_all_test_cases(test_plan):
    tree = ET.parse(test_plan)
    root = tree.getroot()
    count = 0
    step_path = './Testplan/Sequence'
    test_path = step_path + '/Step'
    for item in root.findall(test_path):
        testcase = str(item.find('Test').get('Name'))
        item.set('Disabled', 'True')
        # print(testcase)   # + '- ' + disabled
        count += 1
        # print(str(count) + '. ' + testcase + ' - Disabled')
    tree.write(test_plan)
    return count


if __name__ == '__main__':
    path = input("Enter Report with Test Cases path: ")
    number = input("Enter the Report number (DTP.Execution.xxxx.xlsx): ")
    file = 'DTP.Execution.' + number + '.xlsx'
    if os.path.exists(path):
        aqcc_get_project_TC_Id(path, file)
    else:
        print("No file exists")

    # path = r'E:\Revvl6x Pro'
    # file = 'DTP.Execution.8122.xlsx'
    # if os.path.exists(path):
    #     # aqcc_get_project_TC_Id_complete(path, file)
    #     aqcc_get_project_TC_Id(path, file)
    # else:
    #     print("No file exists")
