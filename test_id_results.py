import os
import re
import pandas as pd
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup

pass_string = "Pass"
fail_string = "Fail"
not_supported = 'Nr CA Combination not supported'

mapping_dict = {}


def textid_mapping():
    path = os.getcwd()
    # df = pd.read_excel(os.path.join(path, "TMO_US_2021_Q4_TestCases.xlsx"), sheet_name='Detail')
    df = pd.read_excel(os.path.join(path, "TMO_US_2023_Q3_TestCases.xlsx"), sheet_name='2023 Q3 Test Cases')
    # tid_tcname_dict = df.set_index('TEST ID')['TEST CASE NAME'].to_dict()
    # Filter 'Subject Path' that starts with 'Protocol'
    df = df.loc[df["Subject Path"].str.startswith('Protocol', na=False)]
    tid_tcname_dict = df.set_index('Id')['Name'].to_dict()
    # print(tid_tcname_dict)
    return tid_tcname_dict


def display_tc_id(testcase_name, file, *args):
    tc_found = False
    verdict = args[0]
    pass_count = args[1]
    ns_count = args[2]
    for key in mapping_dict:
        if testcase_name == str(mapping_dict[key]).lower():
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


def testcaselist(path):
    """
    extract all the testcases in the folder
    """
    count = 0
    pass_count = ns_count = 0
    if os.path.exists(path):
        for file in os.listdir(path):
            if file.startswith('L_'):
                # print(str(count) + "." + file)
                count += 1
                OnlineReport = 'OnlineReport.htm'
                SummaryReport = 'SummaryReport.xml'
                with open("5g_pass_results.txt", "a") as pass_file:
                    with open("5g_result_bcns.txt", "a") as bcns_file:
                        for root, dirs, files in os.walk(path + "\\" + file):
                            verdict = ''
                            if SummaryReport in files:  # and OnlineReport not in files
                                tree = ET.parse(path + "\\" + file + "\\" + SummaryReport)
                                for element in tree.iter(tag='header'):
                                    testcase_number = element.find('testcasenumber').text
                                    observation = element.find('observation').text
                                    final_verdict = element.find('finalverdict').text
                                    # print(str(testcase_number) + '- ' + str(final_verdict) + '- ' + str(observation))
                                    if testcase_number is not None and final_verdict is not None:
                                        testcase_name = testcase_number.lower()
                                        flag = False
                                        if final_verdict == pass_string:
                                            verdict = pass_string
                                            result_file = pass_file
                                            flag = True
                                        elif final_verdict == fail_string and observation == not_supported:  # band Combination is not Supported by Device
                                            verdict = observation
                                            result_file = bcns_file
                                            flag = True
                                        if flag:
                                            tc_found = False
                                            for key in mapping_dict:
                                                if testcase_name == str(mapping_dict[key]).lower():
                                                    tc_found = True
                                                    # print(testcase_name + ' - '+ verdict)
                                                    if verdict == pass_string:
                                                        pass_count += 1
                                                        print(str(pass_count) + ". " + str(key) + " - " + str(
                                                            testcase_name) + " - " + verdict)
                                                    elif verdict == not_supported:
                                                        ns_count += 1
                                                        print(str(ns_count) + ". " + str(key) + " - " + str(
                                                            testcase_name) + " - " + verdict)
                                                    result_file.write(str(key) + "\n")
                                                    break
                                            if not tc_found:
                                                if OnlineReport in files:
                                                    testcase_path = path + "\\" + file + "\\" + OnlineReport
                                                    pass_count = check_in_online_report(testcase_path, pass_count, pass_file)
                                                    # testcase_id, final_verdict, test_case = check_in_online_report(
                                                    # testcase_path) print(str(testcase_id) + ' - ' + str(
                                                    # final_verdict))

        if pass_count > 0 and ns_count == 0:
            print("Total Pass: " + str(pass_count))
        elif pass_count > 0 and ns_count > 0:
            print("Total Pass: " + str(pass_count) + " |\t" + "Total " + not_supported + " : " + str(ns_count))
        elif ns_count > 0:
            print("Total " + not_supported + " : " + str(ns_count))
    else:
        print("Path does not Exists")


def check_in_online_report(OnlineReport, pass_count, pass_file):
    HTMLFile = open(OnlineReport, "r")
    testcase_id = re.findall(r'\bTC-\d+;', HTMLFile.read())[0]
    HTMLFile.close()
    # if not testcase_id:
    testcase_id = testcase_id[:-1]
    HTMLFile = open(OnlineReport, "r")
    testcase_verdict = re.findall(r'\bFinal Verdict:.*Pass.*', HTMLFile.read())
    HTMLFile.close()
    with open(OnlineReport, 'r') as online_report:
        contents = online_report.read()
        soup = BeautifulSoup(contents, 'html.parser')
        div = soup.find("div", {"id": "definitionList_2_section_1"})
        test_case = ''
        for tc in div.find('dd'):
            test_case = tc.text
    if not testcase_verdict:
        final_verdict = "Fail"
    else:
        final_verdict = "Pass"
    if testcase_id is not None and final_verdict == 'Pass':
        pass_count += 1
        pass_file.write(str(testcase_id) + "\n")
        print(
            str(pass_count) + ". " + str(testcase_id) + "- [" + str(test_case) + '] - ' + str(
                final_verdict))
    return pass_count


if __name__ == '__main__':
    if os.path.exists("5g_pass_results.txt"):
        os.remove("5g_pass_results.txt")
    if os.path.exists("5g_result_bcns.txt"):
        os.remove("5g_result_bcns.txt")
    path = input("Enter Result path: ")
    # path = 'D:\Projects\Motorolla\London\Motorola London_2023-02-20T170243'
    print("TestCase in ", path, ":")
    mapping_dict = textid_mapping()
    testcaselist(path)
