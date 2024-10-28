import os
import bs4 as bs


def update_tc_id_list(log_path):
    count = 0
    OverviewReport = 'OverviewReport.htm'
    url = str(log_path + "\\" + OverviewReport)
    file = bs.BeautifulSoup(open(url).read(), "lxml")
    testcase = file.find_all("td", {"class": "tbody_title"})
    with open("lte_contest_pass_results.txt", "a") as pass_file:
        for tc in testcase:
            if tc:
                pass_verdict = tc.parent.find("span", {"class": "verdict_icon pass"})
                if pass_verdict:
                    count += 1
                    print(str(count) + '. '+ tc.get_text() + ' - ' + pass_verdict.get_text())
                    testcase_id = tc.get_text().split(" ", 1)[0]
                    # print(str(testcase_id))
                    pass_file.write(str(testcase_id + "\n"))


if __name__ == '__main__':
    if os.path.exists("lte_contest_pass_results.txt"):
        os.remove("lte_contest_pass_results.txt")
    path = r'D:\Projects\Motorolla\London'
    # path = input("Enter Log path: ")
    update_tc_id_list(path)
