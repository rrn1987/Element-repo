import os
import pandas as pd

mapping_dict = {}


def textid_mapping():
    path = os.getcwd()
    df = pd.read_excel(
        os.path.join(path, "S8706A Protocol Carrier Acceptance Toolset - T-Mobile Test Case Status v14.0.xlsx"),
        sheet_name='Test Case Details', header=3)
    df = df.loc[df['Subject'].str.startswith('Protocol', na=False)]
    tid_tcname_dict = df.set_index('TMO Publication Test ID')['Keysight_ID'].to_dict()
    # print(tid_tcname_dict)
    return tid_tcname_dict


def testcaselist(output_file):
    count = 0
    df = pd.read_excel(output_file, sheet_name='Sheet1')
    # tcname_result_dict = df.set_index('Test Case ID')['Test Result'].to_dict()
    # print(tcname_result_dict)
    for id, row in df.iterrows():
        if row['Test Result'] == 'PASS':
            for key in mapping_dict:
                if row['Test Case ID'] == str(mapping_dict[key]):
                    count += 1
                    print(str(count) + '. ' + str(mapping_dict[key]) + ' - ' + str(key))
                    with open("5g_pass_results.txt", "a") as result_file:
                        result_file.write(str(key) + "\n")
                    break


if __name__ == '__main__':
    if os.path.exists("5g_pass_results.txt"):
        os.remove("5g_pass_results.txt")
    output_file = input("Enter Result file with path: ")
    # output_file = 'D:\Projects\Keysight\Automation\AniteTest_Revel6&6Pro_20230428_183837.xlsx'
    mapping_dict = textid_mapping()
    testcaselist(output_file)
