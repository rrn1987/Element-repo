import os
import pandas as pd
from packaging import version
import xml.etree.ElementTree as ET

NRCA = 'Protocol_NR Carrier Aggregation'
NSA = 'Protocol_5G NSA'
SA = 'Protocol_5G SA'
Emergency_EPS_FB = 'Protocol_5G SA Emergency EPS Fall back'
SA_EPS_FB = 'Protocol_5G SA EPS Fall back'
SA_IMS = 'Protocol_5G SA IMS'
Subject = [SA, NSA, Emergency_EPS_FB, SA_EPS_FB, SA_IMS, NRCA]
# Subject = [NRCA]
Device_Validation_Test_Platform_Complete = 'Device Validation_Test Platform'
Device_Validation_Test_Platform = 'DEVICE VALIDATION TEST PLATFORM'
Platform_Complete = 'Rohde & Schwarz CMX500'
Platform = 'Rohde & Schwarz'
Device_Validation_Category = 'Protocol_5G SA'
RESULT_Remaining = 'NO RUN'
RESULT_Failed = 'FAILED'
versions = []
# OBT_TC_ID = ['TC-19367', 'TC-19377', 'TC-19346', 'TC-19347', 'TC-19345', 'TC-19386', 'TC-19387', 'TC-19385',
# 'TC-19356', 'TC-19357', 'TC-19355', 'TC-19395', 'TC-19396', 'TC-19394', 'TC-19403', 'TC-19404', 'TC-19405',
# 'TC-19816', 'TC-19817', 'TC-19821', 'TC-19822', ' TC-19779', 'TC-19780', 'TC-19775', 'TC-19776', 'TC-19768',
# 'TC-19764', 'TC-19852']

NPT_section = ['Section_Protocol_NR_Carrier_Aggregation', 'Section_Protocol_5G_NSA',
               'Section_Protocol_5G_SA', 'Section_Protocol_5G_SA_VoNR',
               'Section_Protocol_5G_SA_EPS_Fall_back', 'Section_Protocol_5G_NSA_Data_Roaming']
ATE_section = ['Section_Protocol_5G_SA', 'Section_Protocol_5G_NSA', 'Section_Protocol_5G_SA_IMS',
               'Section_Protocol_5G_SA_VoNR',
               'Section_Protocol_5G_SA_EPS_Fall_back', 'Section_Protocol_5G_SA_Emergency_EPS_Fall_back']

# master_path = r'C:\ProgramData\Rohde-Schwarz\Contest\19\Testplans\\'
master_path = r'D:\Projects\Automation\Select TCs in Contest\Testplans\\'
get_version_path = master_path + 'NptTmo'
for filename in os.listdir(get_version_path):
    try:
        # Parse the version number
        ver = version.parse(filename)
        versions.append(ver)
    except version.InvalidVersion:
        print(f"Invalid version format in filename: {filename}")

max_version = max(versions)
print(f"Latest version is: {max_version}")

NPT_path = master_path + 'NptTmo\\' + str(max_version)
ATE_path = master_path + 'AteTmo\\' + str(max_version)


def aqcc_get_project_TC_Id(path, file):
    count = 0
    tc_id_list = []
    df_summary = pd.read_excel(os.path.join(path, file), sheet_name='Summary')
    print("DEVICE INFORMATION: " + df_summary.iat[2, 2])
    print("DEVICE NAME: " + df_summary.iat[3, 2])
    df = pd.read_excel(os.path.join(path, file), sheet_name='Detail')
    # df = df.loc[df[Device_Validation_Test_Platform].str.contains(Platform, na=False)]
    # df = df.loc[df['SUBJECT'].str.startswith((tuple(Subject)), na=False)]
    for id, row in df.iterrows():
        if row['RESULT'] == RESULT_Remaining:
            # print(row['ID'] + ' - ' + row['NAME'] + ' - ' + row['RESULT'])
            count += 1
            tc_id_list.append(row['ID'])
    print('Total TCs required to be selected in the Project for Automation - ' + str(len(tc_id_list)))
    # print(tc_id_list)
    map_AllTestCases_xml(tc_id_list)


def map_AllTestCases_xml(tc_id_list):
    rstt_name_list = []
    file_Alltestcases = os.path.join(os.getcwd(), "Alltestcases.xml")
    tree = ET.parse(file_Alltestcases)
    root = tree.getroot()
    for item in root[0].findall('Testcase'):
        rstt_name = str(item.get('Name'))
        testcase = item.find('Name').text
        if testcase.startswith('TC-'):  # and testcase.count(" ") == 2:
            tc_id = testcase.split(' ')[0]
            testcase_name = testcase.split('-')[2]
            if ' ' in testcase_name:
                testcase_name.replace(' ', '_')
            # print(tc_id)
            # print(testcase_name)
            if tc_id in tc_id_list:
                rstt_name_list.append(rstt_name)
                tc_id_list.remove(tc_id)
    print('Total TCs Selected in Contest - ' + str(len(rstt_name_list)))
    # print(*rstt_name_list)
    # if tc_id_list:
    #     print('Not found in Contest.' + '\n' + str(len(tc_id_list)) + '. ' + str(tc_id_list))
    create_testplan(rstt_name_list)


def create_testplan(rstt_name_list):
    # project_name = input("Enter the Project name: ")
    # shutil.copy(master_file, test_plan)
    each_section = ''
    count = 0
    total = 0
    for each_section in os.listdir(NPT_path):
        test_plan = NPT_path + '\\' + each_section
        disable_count = disable_all_test_cases(test_plan)
        # print('Total TCs Disabled for NPT ' + each_section + ' - ' + str(disable_count))
        count = enable_test_cases(test_plan, rstt_name_list, each_section)
        if count != 0:
            print('Total TCs enabled for NPT ' + each_section + ' - ' + str(count))
        total += count
    for each_section in os.listdir(ATE_path):
        test_plan = ATE_path + '\\' + each_section
        disable_count = disable_all_test_cases(test_plan)
        # print('Total TCs Disabled for ATE ' + each_section + ' - ' + str(disable_count))
        count = enable_test_cases(test_plan, rstt_name_list, each_section)
        if count != 0:
            print('Total TCs enabled for ATE ' + each_section + ' - ' + str(count))
        total += count
    print('Total TCs Enabled:' + str(total))
    # if rstt_name_list:
    #     print('TCs Not Selected: (Removing all SA RTT TCs)')
    #     print([item for item in rstt_name_list if "RTT" not in item])
    #     print(len(rstt_name_list), *rstt_name_list)


def enable_test_cases(test_plan, rstt_name_list, each_section):
    tree = ET.parse(test_plan)
    root = tree.getroot()
    count = 0
    step_path = './Testplan/Sequence'
    test_path = step_path + '/Step'
    for item in root.findall(test_path):
        testcase = str(item.find('Test').get('Name'))
        # print(testcase)   # + '- ' + disabled
        if testcase in rstt_name_list:
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
    if os.path.exists("testcase_name.txt"):
        os.remove("testcase_name.txt")
    # path = input("Enter Report with Test Cases path: ")
    # path = r'C:\Users\instrument\Downloads'
    path = r'D:\Users\nairr\Downloads'
    number = input("Enter the Report number (DTP.Execution.xxxx.xlsx): ")
    file = 'DTP.Execution.' + number + '.xlsx'
    if os.path.exists(path):
        aqcc_get_project_TC_Id(path, file)
    else:
        print("No file exists")

    # path = os.getcwd()
    # file = os.path.join(path, "TMO_US_2023_Q1_TestCases.xlsx")
    # if os.path.exists(path):
    #     aqcc_get_project_TC_Id_complete(path, file)
    # else:
    # print("No file exists")

    # path = 'D:\Projects\Motorolla\Devon'
    # file = 'DTP.Execution.6651.xlsx'
    # if os.path.exists(path):
    #     # aqcc_get_project_TC_Id_complete(path, file)
    #     aqcc_get_project_TC_Id(path, file)
    # else:
    #     print("No file exists")
