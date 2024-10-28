import os
import re
import pandas as pd

testplan = 'TMO_US_2024_Q1_TestCases.xlsx'
sheet_name = 'Test Cases'
Subject = ['Protocol_Carrier Aggregation', 'Protocol_CAT1', 'Protocol_LTE', 'Protocol_Cat-M', 'Protocol_ePDG',
           'Protocol_IMS', 'Protocol_VoLTE', 'Protocol_Wi-Fi Calling', 'Protocol_CBRS_B48']

mapping_dict = {}
pe_mapping_dict = {}
pass_string = "PASS"
fail_string = "FAIL"
bc_ns = "band Combination is not Supported by Device"



# Band Combination is Supported but Layer Combination not Supported by Device

def testcaselist(path):
    """
    extract all the testcaselist in include folder
    """
    total_pass = total_bcns = 0
    pass_count = bcns_count = 0
    if os.path.exists(path):
        for file in os.listdir(path):
            if file.startswith("TestCaseList"):
                testcase_path = path + "\\" + file
                print(testcase_path)
                pass_count, bcns_count = result(testcase_path)
                total_pass += pass_count
                total_bcns += bcns_count
                print("________________________________________________")
        print(
            "Total Test Cases Passed:" + str(total_pass) + "\nTotal Band Combination Not Supported:" + str(total_bcns))
    else:
        print("Path doesn't exists")


def result(file):
    verdict = ''
    L_LTE_ = 'L_LTE_'
    LTE_ = 'LTE_'
    L_IMS_VoLTE_ = 'L_IMS_VoLTE_'
    L_IMS_VOLTE_ = 'L_IMS_VOLTE_'
    L_LTE_CA_ = 'L_LTE_CA_'
    L_CA_ = 'L_CA_'
    L_LAA_ = 'L_LAA_'
    pass_count = bcns_count = 0
    with open(file) as file:
        with open("lte_pe_result_pass.txt", "a") as pass_file:
            with open("lte_pe_result_bcns.txt", "a") as bcns_file:
                for line in file:
                    if line is not None and line.startswith(r'document.write("<table'):
                        if pass_string in line or (fail_string in line and bc_ns in line):
                            if pass_string in line:
                                verdict = pass_string
                                result_file = pass_file
                            elif fail_string in line and bc_ns in line:  # band Combination is not Supported by Device
                                verdict = bc_ns
                                result_file = bcns_file
                            if L_LTE_ in line or LTE_ in line and L_IMS_VoLTE_ not in line and L_IMS_VOLTE_ not in line:
                                testcase_name, flag = check_specific(line)
                                if not flag:
                                    testcase_name = re.findall(r'\bL_LTE_([.\w]+)', repr(line)) \
                                                    or re.findall(r'\bLTE_([.\w]+)', repr(line))
                                    # print('L_LTE_' + testcase_name[0])
                                    testcase_name = 'L_LTE_' + testcase_name[0]
                                pass_count, bcns_count = display_tc_id(testcase_name, result_file, 'CA', verdict,
                                                                       pass_count, bcns_count)
                            elif L_IMS_VoLTE_ in line or L_IMS_VOLTE_ in line:  # For VOLTE Results
                                testcase_name = re.findall(r'\bL_IMS_VoLTE_([.\w]+)', repr(line)) \
                                                or re.findall(r'\bL_IMS_VOLTE_([.\w]+)', repr(line))
                                # print('{0} : {1}'.format(count, 'L_IMS_VoLTE_' + testcase_name[0]))
                                testcase_name = 'L_IMS_VoLTE_' + testcase_name[0]
                                pass_count, bcns_count = display_tc_id(testcase_name, result_file, 'VoLTE', verdict,
                                                                       pass_count, bcns_count)
                            elif L_LTE_CA_ in line:  # TM9
                                testcase_name = re.findall(r'\bL_LTE_CA_([.\w]+)', repr(line))
                                # print('{0} : {1}'.format(count, 'L_LTE_' + testcase_name[0]))
                                testcase_name = 'L_LTE_CA_' + testcase_name[0]
                                pass_count, bcns_count = display_tc_id(testcase_name, result_file, 'CA', verdict,
                                                                       pass_count, bcns_count)
                            elif L_CA_ in line:  # B41CA+MIMO
                                testcase_name, flag = check_specific(line)
                                if not flag:
                                    testcase_name = re.findall(r'\bL_CA_([.\w]+)', repr(line))
                                    # print('{0} : {1}'.format(count, 'L_LTE_' + testcase_name[0]))
                                    testcase_name = 'L_CA_' + testcase_name[0]
                                pass_count, bcns_count = display_tc_id(testcase_name, result_file, 'CA', verdict,
                                                                       pass_count, bcns_count)
                            elif L_LAA_ in line:
                                testcase_name = re.findall(r'\bL_LAA_([.\w]+)', repr(line))
                                # print('{0} : {1}'.format(count, 'L_LTE_' + testcase_name[0]))
                                testcase_name = 'L_LAA_' + testcase_name[0]
                                pass_count, bcns_count = display_tc_id(testcase_name, result_file, 'LAA', verdict,
                                                                       pass_count, bcns_count)
                            else:
                                print("Fail")
            if pass_count > 0 and bcns_count == 0:
                print("Total Pass in Section: " + str(pass_count))
            elif pass_count > 0 and bcns_count > 0:
                print("Total Pass in Section: " + str(pass_count) + "\t" + "Total BC NS in Section: " + str(bcns_count))
            elif bcns_count > 0:
                print("Total BC NS in Section: " + str(bcns_count))
            return pass_count, bcns_count


def check_specific(line):
    testcase_name = ''
    flag = False
    if 'L_01_218202_LTE_UL256QAM' in line:
        testcase_name = 'L_01_218202_LTE_UL256QAM'
        flag = True
    elif 'L_CA_513616_TM mode switch_B41' in line:
        testcase_name = 'L_CA_513616_TM mode switch_B41'
        flag = True
    elif 'L_CA_513617_1_4x4_TM3 MIMO B41 VoLTE' in line:
        testcase_name = 'L_CA_513617_1_4x4_TM3 MIMO B41 VoLTE'
        flag = True
    elif 'L_CA_513617_1_4x4_TM3 MIMO B41' in line:
        testcase_name = 'L_CA_513617_1_4x4_TM3 MIMO B41'
        flag = True
    elif 'L_CA_513617_1_TM8 MIMO B41 VoLTE' in line:
        testcase_name = 'L_CA_513617_1_TM8 MIMO B41 VoLTE'
        flag = True
    elif 'L_CA_513617_1_TM8 MIMO B41' in line:
        testcase_name = 'L_CA_513617_1_TM8 MIMO B41'
        flag = True
    elif 'L_CA_513619_1_ULCA [41C UL + 41D DL]' in line:
        testcase_name = 'L_CA_513619_1_ULCA [41C UL + 41D DL]'
        flag = True
    elif 'L_CA_513619_1_ULCA [41C]' in line:
        testcase_name = 'L_CA_513619_1_ULCA [41C]'
        flag = True
    elif 'L_LTE_98201_2_TM9_LTE B2' in line:
        testcase_name = 'L_LTE_98201_2_TM9_LTE B2'
        flag = True
    elif 'L_LTE_98201_2_TM9_LTE B4' in line:
        testcase_name = 'L_LTE_98201_2_TM9_LTE B4'
        flag = True
    elif 'L_LTE_98201_2_TM9_LTE B66' in line:
        testcase_name = 'L_LTE_98201_2_TM9_LTE B66'
        flag = True
    elif 'L_LTE_98201_1_TM9_4x4_MIMO_LTE_B4' in line:
        testcase_name = 'L_LTE_98201_1'
        flag = True
    elif 'L_LTE_175002_01' in line:
        testcase_name = 'L_LTE_175002_1'
        flag = True
    elif 'L_LTE_175002_02' in line:
        testcase_name = 'L_LTE_175002_2'
        flag = True
    elif 'L_LTE_175002_03' in line:
        testcase_name = 'L_LTE_175002_3'
        flag = True
    elif 'L_01_519371_LTE_Band25' in line:
        testcase_name = 'L_01_519371_LTE_Band25 LTE Band 25 Attach, VoLTE call, data streaming'
        flag = True
    elif 'L_01_519372_LTE_Band26' in line:
        testcase_name = 'L_01_519372_LTE_Band26 LTE Band 26 Attach, VoLTE call, data streaming'
        flag = True
    elif 'L_02_519371_LTE_Band25' in line:
        testcase_name = 'L_02_519371_LTE_Band25 VoLTE call and PS HO_LTE Band 25 to LTE NSN Band 4,5,12,66,41'
        flag = True
    elif 'L_02_519372_LTE_Band26' in line:
        testcase_name = 'L_02_519372_LTE_Band26 VoLTE call and PS HO_LTE Band 26 to LTE NSN Band 4,5,12,66,41'
        flag = True
    return testcase_name, flag


def display_tc_id(testcase_name, file, *args):
    tc_found = False
    section = args[0]
    verdict = args[1]
    pass_count = args[2]
    bcns_count = args[3]
    for key in mapping_dict:
        if testcase_name == str(mapping_dict[key]) or (section == 'CA' and testcase_name in str(mapping_dict[key])):
            tc_found = True
            if verdict == pass_string:
                pass_count += 1
                print(str(pass_count) + ": " + str(key) + ":" + str(testcase_name) + " - " + verdict)
            elif verdict == bc_ns:
                bcns_count += 1
                print(str(bcns_count) + ": " + str(key) + ":" + str(testcase_name) + " - " + verdict)
            file.write(str(key) + "\n")
            break
    for pe_key in pe_mapping_dict:
        if testcase_name == str(pe_mapping_dict[pe_key]):
            tc_found = True
            if verdict == pass_string:
                pass_count += 1
                print(str(pass_count) + ": " + str(pe_key) + ":" + str(testcase_name) + " - " + verdict)
            elif verdict == bc_ns:
                bcns_count += 1
                print(str(bcns_count) + ": " + str(pe_key) + ":" + str(testcase_name) + " - " + verdict)
            file.write(str(pe_key) + "\n")
            break
    if not tc_found:
        print(str(testcase_name) + " Not found in the TMO 2023 Q4 Sheet or Deleted TC")
    return pass_count, bcns_count


def textid_mapping():
    tid_tcname_dict = {}
    path = os.getcwd()
    df = pd.read_excel(os.path.join(path, testplan), sheet_name=sheet_name)
    # Filter 'Subject Path' that starts with 'Protocol'
    df = df.loc[df["Subject Path"].str.startswith((tuple(Subject)), na=False)]
    tid_tcname_dict = df.set_index('Id')['Name'].to_dict()
    replace_dict_testcase_name(tid_tcname_dict)
    # print(tid_tcname_dict)
    return tid_tcname_dict

def pe_textid_mapping():
    path = os.getcwd()
    df = pd.read_excel(os.path.join(path, "LTE PE TC ID Mapping.xlsx"), sheet_name='LTE PE ID')
    # Filter 'Subject Path' that starts with 'Protocol'
    df = df.loc[df["Subject Path"].str.startswith('Protocol', na=False)]
    pe_tcid_tcname_dict = df.set_index('Id')['PE ID'].to_dict()
    # print(pe_tcid_tcname_dict)
    return pe_tcid_tcname_dict


def replace_dict_testcase_name(tid_tcname_dict):
    for key in tid_tcname_dict:
        if tid_tcname_dict[key] == 'L_LTE_380901_1_UE capability_8x2 MIMO':
            tid_tcname_dict[key] = 'L_LTE_380901_1'
        elif tid_tcname_dict[key] == 'L_LTE_380901_2_8x2 MIMO_PS_LTE B4':
            tid_tcname_dict[key] = 'L_LTE_380901_2'
        elif tid_tcname_dict[key] == 'L_LTE_380901_3_8x4 MIMO_PS_LTE B4':
            tid_tcname_dict[key] = 'L_LTE_380901_3'
        elif tid_tcname_dict[key] == 'L_LTE_380901_4_8x2 MIMO_PS_LTE B71':
            tid_tcname_dict[key] = 'L_LTE_380901_4'
        elif tid_tcname_dict[key] == 'L_LTE_380901_5_8x4 MIMO_PS_LTE B71':
            tid_tcname_dict[key] = 'L_LTE_380901_5'
        elif tid_tcname_dict[key] == 'L_LTE_380901_6_8x2 MIMO_PS_LTE B66':
            tid_tcname_dict[key] = 'L_LTE_380901_6'
        elif tid_tcname_dict[key] == 'L_LTE_380901_7_8x4 MIMO_PS_LTE B66':
            tid_tcname_dict[key] = 'L_LTE_380901_7'
        elif tid_tcname_dict[key] == 'L_LTE_380901_8_8x2 MIMO_VoLTE emergency call':
            tid_tcname_dict[key] = 'L_LTE_380901_8'
        elif tid_tcname_dict[key] == 'L_LTE_380901_9_8x4 MIMO_VoLTE emergency call':
            tid_tcname_dict[key] = 'L_LTE_380901_9'
        elif tid_tcname_dict[key] == 'L_LTE_380901_10_Mobility between 8x2 and 4x2':
            tid_tcname_dict[key] = 'L_LTE_380901_10'
        elif tid_tcname_dict[key] == 'L_LTE_380901_11_Mobility between 8x4 and 4x4':
            tid_tcname_dict[key] = 'L_LTE_380901_11'
        elif tid_tcname_dict[key] == 'L_LTE_7312_Downlink 4x2 MIMO shall be supported for all LTE B4':
            tid_tcname_dict[key] = 'L_LTE_7312_1'
        elif tid_tcname_dict[key] == 'L_LTE_7312_Downlink 4x2 MIMO shall be supported for all LTE B2':
            tid_tcname_dict[key] = 'L_LTE_7312_2'
        elif tid_tcname_dict[key] == 'L_LTE_7312_Downlink 4x2 MIMO shall be supported for all LTE B66':
            tid_tcname_dict[key] = 'L_LTE_7312_3'
        elif tid_tcname_dict[key] == 'L_LTE_480562_HO TM9 to TM9':
            tid_tcname_dict[key] = 'L_LTE_480562_HO_TM9toTM9'
        elif tid_tcname_dict[key] == 'L_LTE_480562_CA_B2+B66_TM9+TM4':
            tid_tcname_dict[key] = 'L_LTE_480562_CA_B2_B66_TM9_TM4'
        elif tid_tcname_dict[key] == 'L_LTE_480562_CA_B2+B66_TM9':
            tid_tcname_dict[key] = 'L_LTE_480562_CA_B2_B66_TM9'
        elif tid_tcname_dict[key] == 'L_LTE_480562_TM mode switch_B2':
            tid_tcname_dict[key] = 'L_LTE_480562_TM_Mode_Switch_B2'
        elif tid_tcname_dict[key] == 'L_LTE_480562_TM mode switch_B66':
            tid_tcname_dict[key] = 'L_LTE_480562_TM_Mode_Switch_B66'
        elif tid_tcname_dict[key] == 'L_LTE_460434_1_TM9 4x2 MIMO_256QAM':
            tid_tcname_dict[key] = 'L_LTE_460434_1_TM9_4x2_MIMO'
        elif tid_tcname_dict[key] == 'L_LTE_460434_2_TM9 4x4 MIMO_256QAM':
            tid_tcname_dict[key] = 'L_LTE_460434_2_TM9_4x4_MIMO'
        elif tid_tcname_dict[key] == 'L_LTE_460434_3_TM9 8x2 MIMO_256QAM':
            tid_tcname_dict[key] = 'L_LTE_460434_3_TM9_8x2_MIMO'
        elif tid_tcname_dict[key] == 'L_LTE_460434_4_TM9 8x4 MIMO_256QAM':
            tid_tcname_dict[key] = 'L_LTE_460434_4_TM9_8x4_MIMO'
        elif 'L_LTE_7201_2' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_7201_2'
        elif 'L_LTE_7147' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_7147'
        elif 'L_LTE_7239' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_7239'
        elif 'L_LTE_218201_5' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218201_5'
        elif 'L_LTE_218201_6' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218201_6'
        elif 'L_LTE_218201_7' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218201_7'
        elif 'L_LTE_218201_10' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218201_10'
        elif 'L_LTE_218201_12' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218201_12'
        elif 'L_LTE_218201_1' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218201_1'
        elif 'L_LTE_218202_5' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218202_5'
        elif 'L_LTE_218202_6' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218202_6'
        elif 'L_LTE_218202_7' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218202_7'
        elif 'L_LTE_218202_8' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218202_8'
        elif 'L_LTE_218202_1' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_LTE_218202_1'
        elif 'L_IMS_VoLTE_6825_1' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = 'L_IMS_VoLTE_6825_1'
        elif 'L_LAA' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = str(tid_tcname_dict[key]).replace(" ", "")
        elif 'L_LTE_' in str(tid_tcname_dict[key]):
            tid_tcname_dict[key] = str(tid_tcname_dict[key]).replace(" ", "")
    # print(tid_tcname_dict)


if __name__ == '__main__':
    # file_path = input("Enter the include path:")
    if os.path.exists("lte_pe_result_pass.txt"):
        os.remove("lte_pe_result_pass.txt")
    if os.path.exists("lte_pe_result_bcns.txt"):
        os.remove("lte_pe_result_bcns.txt")
    path = input("Enter Result path: ")
    path = path + "\\include"
    # path = r'D:\Projects\Motorolla\London\LTE\Run1\include'
    print("TestCaseList in ", path, ":")
    mapping_dict = textid_mapping()
    pe_mapping_dict = pe_textid_mapping()
    testcaselist(path)
