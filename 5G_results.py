import os
import re
import pandas as pd
import xml.etree.ElementTree as ET

mapping_dict = {}
pass_string = "Pass"
fail_string = "Fail"
not_supported = 'Nr CA Combination not supported'


def textid_mapping():
    tid_tcname_dict = {}
    path = os.getcwd()
    # df = pd.read_excel(os.path.join(path, "TMO_US_2021_Q4_TestCases.xlsx"), sheet_name='Detail')
    df = pd.read_excel(os.path.join(path, "TMO_US_2023_Q1_TestCases.xlsx"), sheet_name='2023Q1 Test Cases')
    # tid_tcname_dict = df.set_index('TEST ID')['TEST CASE NAME'].to_dict()
    # Filter 'Subject Path' that starts with 'Protocol'
    df = df.loc[df["Subject Path"].str.startswith('Protocol_NR Carrier Aggregation', na=False)]
    tid_tcname_dict = df.set_index('Id')['Name'].to_dict()
    # replace_dict_testcase_name(tid_tcname_dict)
    # print(tid_tcname_dict)
    return tid_tcname_dict


def display_tc_id(testcase_name, file, *args):
    tc_found = False
    section = args[0]
    verdict = args[1]
    pass_count = args[2]
    ns_count = args[3]
    for key in mapping_dict:
        if testcase_name == str(mapping_dict[key]):
            tc_found = True
            # print(testcase_name + ' - '+ verdict)
            if verdict == pass_string:
                pass_count += 1
                print(str(pass_count) + ": " + str(key) + ":" + str(testcase_name) + " - " + verdict)
            elif verdict == not_supported:
                ns_count += 1
                print(str(ns_count) + ": " + str(key) + ":" + str(testcase_name) + " - " + verdict)
            file.write(str(key) + "\n")
            break
    if not tc_found:
        print(str(testcase_name) + " Not found in the TMO 2023 Q1 Sheet")
    return pass_count, ns_count


if __name__ == '__main__':
    if os.path.exists("5g_pass_results.txt"):
        os.remove("5g_pass_results.txt")
    if os.path.exists("5g_result_bcns.txt"):
        os.remove("5g_result_bcns.txt")
    # path = input("Enter Result path: ")
    path = r'D:\Projects\Automation'
    file = path + '\\SummaryReportsOverview.xml'
    # mapping_dict = textid_mapping()
    mytree = ET.parse(file)
    myroot = mytree.getroot()
    pass_count = ns_count = 0
    L_NRCA = 'L_NRCA_'
    FiveG_ = '5G_'
    with open("5g_pass_results.txt", "a") as pass_file:
        with open("5g_result_bcns.txt", "a") as bcns_file:
            for x in myroot[1].findall('testreport'):
                verdict = result = ''
                testcase = x.find('testcase').text
                observation = x.find('observation').text
                final_verdict = x.find('finalverdict')
                result = str(final_verdict.get('result'))
                if testcase is not None and result is not None:
                    if testcase.startswith(L_NRCA):
                        flag = False
                        if result == pass_string:
                            verdict = pass_string
                            result_file = pass_file
                            flag = True
                        elif result == fail_string and observation == not_supported:  # band Combination is not Supported by Device
                            verdict = not_supported
                            result_file = bcns_file
                            flag = True
                        if flag:
                            testcase_name = re.match(r'\bL_NRCA_.*\d+\.\d+\s', testcase)
                            testcase_name = testcase_name.group(0).strip()
                            pass_count, ns_count = display_tc_id(testcase_name, result_file, 'L_NRCA_', verdict, pass_count, ns_count)
                        # else:
                        #     testcase_name = re.findall(r'\b5G_([.\w]+)', repr(testcase)) \
                        #                     or re.findall(r'\bENDC_IMS_([.\w]+)', repr(testcase))
                        #     testcase_name = '5G_' + testcase_name[0]
                        #     pass_count = display_tc_id(testcase_name, result_file, '5G', result, pass_count)
                else:
                    print(str(pass_count) + ". TestCase: " + str(testcase) + "\nResult is None")
                    pass_count += 1
            if pass_count > 0 and ns_count == 0:
                print("Total Pass: " + str(pass_count))
            elif pass_count > 0 and ns_count > 0:
                print("Total Pass: " + str(pass_count) + "\t" + "Total BC Not Supported: " + str(ns_count))
            elif ns_count > 0:
                print("Total BC Not Supported: " + str(ns_count))
