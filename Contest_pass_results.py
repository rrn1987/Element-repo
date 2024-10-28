import os
import bs4 as bs


def update_tc_id_list(log_path):
    count = 0
    testcase_id = ''
    OverviewReport = 'OverviewReport.htm'
    url = str(log_path + "\\" + OverviewReport)
    with open(url, 'rb') as html:
        file = bs.BeautifulSoup(html)
        # file = bs.BeautifulSoup(open(url).read(), "lxml", errors="ignore")
        testcase = file.find_all("td", {"class": "tbody_title"})
        with open("contest_pass_results.txt", "a") as pass_file:
            for tc in testcase:
                if tc:
                    pass_verdict = tc.parent.find("span", {"class": "verdict_icon pass"})
                    if pass_verdict:
                        count += 1
                        print(str(count) + '. ' + tc.get_text() + ' - ' + pass_verdict.get_text())
                        if tc.get_text().split(" ")[2].startswith('TC-'):  # NetOp TCID
                            testcase_id = tc.get_text().split(" ")[2]
                        elif tc.get_text().split(" ")[0].startswith('TC-'):  # ATE TCID
                            testcase_id = tc.get_text().split(" ")[0]
                        # print(str(count) + '. ' + str(testcase_id))
                        pass_file.write(str(testcase_id + "\n"))


if __name__ == '__main__':
    if os.path.exists("contest_pass_results.txt"):
        os.remove("contest_pass_results.txt")
    path = input("Enter Log path: ")
    update_tc_id_list(path)
