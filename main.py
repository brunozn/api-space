from datetime import datetime
import sys
import requests
import os
import xlsxwriter

# from datetime import datetime
TOKEN = os.getenv('TOKEN')
URL = 'https://vortexz.jetbrains.space/api/http/projects/key:EVENLINEZ/planning/issues?sorting=CREATED&' \
      'descending=true&$fields=data(status(name,id),title,sprints(board(board(id,name)),description),' \
      'creationTime,description)'


def get_issues(sprint=None):
    headers = {'Authorization': 'Bearer ' + TOKEN, 'Accept': 'application/json'}
    datas = requests.get(URL, headers=headers).json()["data"]
    issues = []
    for element in datas:
        if sprint:
            for y in element["sprints"]:
                if y["board"]["board"]["name"] == str(sprint):
                    issue = {
                        "Title": element["title"],
                        "Description": element["description"],
                        "Created": element["creationTime"]["iso"],
                        "Status": element["status"]["name"],
                        "Sprint": y["board"]["board"]["name"]
                    }
                    issues.append(issue)
        else:
            issue = {
                "Title": element["title"],
                "Description": element["description"],
                "Created": element["creationTime"]["iso"],
                "Status": element["status"]["name"],
                "Sprint": None
            }
            issues.append(issue)
    return issues


def export_excel(isus):
    book = xlsxwriter.Workbook('Issues.xlsx')
    sheet = book.add_worksheet("Issues BOARD")
    sheet.set_column('A:A', 60)
    sheet.set_column('B:B', 50)
    sheet.set_column('C:C', 15)
    sheet.set_column('D:D', 15)
    sheet.merge_range('A1:F1', "SPRINT", book.add_format({'bold': True, 'font_size': 14, 'align': 'center',
                                                          'valign': 'vcenter', 'font_name': 'Arial',
                                                          'color': '#171717'}))
    bold = book.add_format({'bold': 1})
    row_num = 1
    columns = ['Title', 'Description', 'Created', 'Status']
    for col_num in range(len(columns)):
        sheet.write(row_num, col_num, columns[col_num], bold)

    for idy, data in enumerate(isus):
        date_format = datetime.strftime(datetime.strptime((data['Created']), "%Y-%m-%dT%H:%M:%S.%fZ"), "%d/%m/%Y")
        col = 0
        row = 2 + idy
        if data['Title'] is None:
            sheet.write_string(row, col, '', )
        else:
            sheet.write_string(row, col, data['Title'])
        if data['Description'] is None:
            sheet.write_string(row, col + 1, '', )
        else:
            sheet.write_string(row, col + 1, data['Description'])
        sheet.write(row, col + 2, date_format)
        sheet.write_string(row, col + 3, data['Status'])
        if data['Sprint'] is None:
            sheet.write_string(row, col + 4, '')
        else:
            sheet.write_string(row, col + 4, data['Sprint'])
    book.close()


def main(ar):
    arg = None
    if len(ar) > 1: arg = sys.argv[1]
    issues = get_issues(arg)
    export_excel(issues)


main(sys.argv)
