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
    if sprint:
        sprint = str(sprint)
        for x in datas:
            for y in x["sprints"]:
                if y["board"]["board"]["name"] == sprint:
                    issue = {
                        "Title": x["title"],
                        "Description": x["description"],
                        "Created": x["creationTime"]["iso"],
                        "Status": x["status"]["name"],
                    }
                    issues.append(issue)
    else:
        for element in datas:
            issue = {
                "Title": element["title"],
                "Description": element["description"],
                "Created": element["creationTime"]["iso"],
                "Status": element["status"]["name"],
            }
            issues.append(issue)
    return issues


def export_excel(isus):
    book = xlsxwriter.Workbook('Issues.xlsx')
    sheet = book.add_worksheet("Issues BOARD")
    sheet.set_column('A:A', 60)
    sheet.set_column('B:B', 40)
    sheet.set_column('C:C', 30)
    sheet.merge_range('A1:F1', "SPRINT", book.add_format({'bold': True, 'font_size': 14, 'align': 'center',
                                                          'valign': 'vcenter', 'font_name': 'Arial',
                                                          'color': '#171717'}))
    bold = book.add_format({'bold': 1})
    row_num = 1
    columns = ['Title', 'Description', 'Created', 'Status']
    for col_num in range(len(columns)):
        sheet.write(row_num, col_num, columns[col_num], bold)

    for idy, data in enumerate(isus):
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
        sheet.write_string(row, col + 2, data['Created'])
        sheet.write_string(row, col + 3, data['Status'])
    book.close()


issues = get_issues()
export_excel(issues)
