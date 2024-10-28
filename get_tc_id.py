import os
import re
import xml.etree.ElementTree as ET


pass_string = "Pass"
fail_string = "Fail"
not_supported = 'Nr CA Combination not supported'
SummaryReportsOverview = 'SummaryReportsOverview.xml'
testcase_id_list = []


def get_tc_id_list(path):
    """
    extract all the testcase id in the SummaryReportsOverview
    """
    pass_count = ns_count = 0
    if os.path.exists(path):
        file = path + '\\' + SummaryReportsOverview
        mytree = ET.parse(file)
        myroot = mytree.getroot()
        pass_count = ns_count = 0
        with open("5g_pass_results.txt", "a") as pass_file:
            with open("5g_result_bcns.txt", "a") as bcns_file:
                for x in myroot[1].findall('testreport'):
                    verdict = ''
                    testcase = x.find('testcase').text
                    observation = x.find('observation').text
                    final_verdict = x.find('finalverdict')
                    result = str(final_verdict.get('result'))
                    # print(str(testcase) + ' - ' + str(result) + ' - ' + str(observation))
                    if testcase is not None and result is not None:
                        flag = False
                        if result == pass_string:
                            verdict = pass_string
                            result_file = pass_file
                            flag = True
                            # print(str(testcase) + ' - ' + str(result) + ' - ' + str(observation))
                        elif result == fail_string and observation == not_supported:  # band Combination is not Supported by Device
                            verdict = observation
                            result_file = bcns_file
                            flag = True
                        if flag:
                            testcase_id = re.findall(r'\bTC-\d+', testcase)[0]
                            testcase_id_list.append(testcase_id)
                            # print(str(testcase_id) + ' - ' + verdict)
                            if verdict == pass_string:
                                pass_count += 1
                                print(str(pass_count) + ". " + str(testcase) + " - " + verdict)
                            elif verdict == not_supported:
                                ns_count += 1
                                print(str(ns_count) + ". " + str(testcase) + " - " + verdict)
                            result_file.write(str(testcase_id) + "\n")

        # print(testcase_id_list)
        if pass_count > 0 and ns_count == 0:
            print("Total Pass: " + str(pass_count))
        elif pass_count > 0 and ns_count > 0:
            print("Total Pass: " + str(pass_count) + " |\t" + "Total " + not_supported + " : " + str(ns_count))
        elif ns_count > 0:
            print("Total " + not_supported + " : " + str(ns_count))
    else:
        print("Path does not Exists")


if __name__ == '__main__':
    if os.path.exists("5g_pass_results.txt"):
        os.remove("5g_pass_results.txt")
    if os.path.exists("5g_result_bcns.txt"):
        os.remove("5g_result_bcns.txt")
    path = input("Enter Result path: ")
    # path = 'D:\Projects\Automation'
    print("TestCase in ", path, ":")
    get_tc_id_list(path)
