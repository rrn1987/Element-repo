import os
from builtins import reversed

import pandas as pd
from collections import OrderedDict

file_name = 'DTP.Execution.8561.xlsx'


def group_list(lst):
    res = [(el, lst.count(el)) for el in lst]
    return list((OrderedDict(res).items()))


def textid_mapping():
    list_pcom_id = []
    list_doc_id = []
    count = 0
    source_path = r'D:\Projects\ATT\10776 Test Scripts - 23.3.1'
    source_df = pd.read_excel(os.path.join(source_path, "10776 23.3.1.xlsx"), sheet_name='Test Cases', header=6,
                              keep_default_na=False)
    tcname_pcom_dict = source_df.set_index('Test case Name')['PCOM Script ID'].to_dict()
    tcname_docid_dict  = source_df.set_index('Test case Name')['AT&T Test Document Number'].to_dict()
    project_path = r'D:\Projects\ATT'
    project_df = pd.read_excel(os.path.join(project_path, file_name), sheet_name='Execution Results')
    for tc_name in project_df['TCID']:
        for key in tcname_pcom_dict:
            if tc_name == str(key):
                count += 1
                # print(str(count) + '.' + str(tc_name) + ' - ' + str(tcname_pcom_dict[key]))
                list_pcom_id.append(str(tcname_pcom_dict[key]))
                list_doc_id.append(str(tcname_docid_dict[key]))
                break
            else:
                if tc_name == 'LTE-BTR-5-5400.6':
                    count += 1
                    print(str(count) + '.' + str(tc_name) + ' Not in 10776 sheet')
                    tcname_pcom_dict[key] = 'Not in 10776 sheet'
                    tcname_docid_dict[key] = 'Not in 10776 sheet'
                    list_pcom_id.append(str(tcname_pcom_dict[key]))
                    list_doc_id.append(str(tcname_docid_dict[key]))
                    break
    # print(len(list_pcom_id))
    # print(len(list_doc_id))
    project_df.insert(4, 'PCOM ID', list_pcom_id, True)
    project_df.insert(5, 'AT&T Test Document Number', list_doc_id, True)
    pd.set_option('display.max_rows', None)
    # print(project_df[['TCID', 'PCOM ID']])
    writer = pd.ExcelWriter(project_path + '/OutputFile.xlsx', engine='xlsxwriter')
    project_df.to_excel(writer, sheet_name='Test Case', index=1,)
    for column in project_df:
        col_x = project_df.columns.get_loc('PCOM ID')
        writer.sheets['Test Case'].set_column(col_x, col_x, 30)
        col_idx = project_df.columns.get_loc('AT&T Test Document Number')
        writer.sheets['Test Case'].set_column(col_idx, col_idx, 30)
    writer.save()
    print('Written to Excel File successfully.')
    print(group_list(list_pcom_id))


if __name__ == '__main__':
    path = r'D:\Projects\ATT'
    textid_mapping()
