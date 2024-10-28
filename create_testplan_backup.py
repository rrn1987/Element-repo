import os
import pandas as pd
import shutil
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


def aqcc_testlist_complete_byId(path, file):
    count = 0
    testcase_list = []
    df = pd.read_excel(os.path.join(path, file), sheet_name='2023Q1 Test Cases')
    df = df.loc[df["Device Validation_Test Platform"].str.contains(Platform_Complete, na=False)]
    df = df.loc[df["Subject Path"].str.startswith((tuple(Subject)), na=False)]
    # if Subject[0] == SA:
    #     df = df.loc[df["Device Validation_Category"] == SA]
    for Id in df['Id']:
        count += 1
        testcase_list.append(Id)
    print('Total TCs in Q1 Test Plan - ' + str(len(testcase_list)))
    map_Id_TestCase(testcase_list)


def map_Id_TestCase(testcase_list):
    tc_id_list = []
    count = 0
    df = pd.read_excel(os.path.join(os.getcwd(), 'TMO_NPT_ATE_PQA_Q1_2023_Release.xlsx'), sheet_name='Data')
    tid_tcname_dict = df.set_index('Test ID')['Test Case'].to_dict()
    for Id in testcase_list:
        if Id in tid_tcname_dict:
            count += 1
            # print(str(count) + ' - ' + Id + ' - ' + str(tid_tcname_dict[Id]))
            tc_id_list.append(str(tid_tcname_dict[Id]))
        else:
            print(Id + ' - Not Found in R&S Sheet')
    print('Total TCs Selected in ATE NPT Sheet - ' + str(len(tc_id_list)))
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
        if testcase in tc_id_list:
            rstt_name_list.append(rstt_name)
            tc_id_list.remove(testcase)
    print('Total TCs Selected in XML File - ' + str(len(rstt_name_list)))
    if tc_id_list:
        print(str(*tc_id_list) + ' - Not found in XML File')
    create_testplan(rstt_name_list)


def create_testplan(rstt_name_list):
    # project_name = input("Enter the Project name: ")
    # shutil.copy(master_file, test_plan)

    global each_section, count
    NPT_section = ['Section_Protocol_NR_Carrier_Aggregation.rstt', 'Section_Protocol_5G_NSA.rstt',
                   'Section_Protocol_5G_SA.rstt', 'Section_Protocol_5G_SA_VoNR.rstt',
                   'Section_Protocol_5G_SA_EPS_Fall_back.rstt']
    ATE_section = ['Section_Protocol_5G_NSA.rstt', 'Section_Protocol_5G_SA_IMS.rstt', 'Section_Protocol_5G_SA_VoNR'
                                                                                      '.rstt',
                   'Section_Protocol_5G_SA_EPS_Fall_back.rstt', 'Section_Protocol_5G_SA_Emergency_EPS_Fall_back.rstt']

    # master_path = r'C:\ProgramData\Rohde-Schwarz\Contest\19\Testplans\\'
    master_path = r'D:\Projects\Automation\Select TCs in Contest\Testplans\\'
    NPT_path = master_path + 'NptTmo\9.2'
    ATE_path = master_path + 'AteTmo\9.2'
    total = 0
    for each_section in NPT_section:
        # print('NPT Section:' + each_section)
        test_plan = NPT_path + '\\' + each_section
        count = enable_test_cases(test_plan, rstt_name_list, each_section)
        print('Total TCs enabled for ' + each_section + ' - ' + str(count))
        total += count
    for each_section in ATE_section:
        # print('ATE Section:' + each_section)
        test_plan = ATE_path + '\\' + each_section
        count = enable_test_cases(test_plan, rstt_name_list, each_section)
        print('Total TCs enabled for ' + each_section + ' - ' + str(count))
        total += count
    print('Total TCs Enabled:' + str(total))
    if rstt_name_list:
        print(*rstt_name_list)

    # print(len(rstt_name_list), *rstt_name_list)


def enable_test_cases(test_plan, rstt_name_list, each_section):
    tree = ET.parse(test_plan)
    root = tree.getroot()
    count = 0
    step_path = './Testplan/Sequence'
    test_path = step_path + '/Step'
    with open("testcase_name.txt", "a") as testcase_name_file:
        for item in root.findall(test_path):
            id = str(item.get('ID'))
            testcase = str(item.find('Test').get('Name'))
            testcase_name_file.write(str(id) + '\t' + testcase + '\n')
            # print(testcase)   # + '- ' + disabled
            if testcase in rstt_name_list:
                count += 1
                rstt_name_list.remove(testcase)
                # print(str(count) + '. ' + testcase + ' - Enabled')
                item.set('Disabled', 'False')
    # tree.write(test_plan)
    return count


if __name__ == '__main__':
    if os.path.exists("testcase_name.txt"):
        os.remove("testcase_name.txt")
    # path = input("Enter Report Sheet with Test Cases path: ")
    path = os.getcwd()
    file = os.path.join(path, "TMO_US_2023_Q1_TestCases.xlsx")
    # path = 'D:\Device\XT2305-1(London)'
    # file = 'DTP.Execution.6621.xlsx'
    if os.path.exists(path):
        aqcc_testlist_complete_byId(path, file)
    else:
        print("No file exists")
