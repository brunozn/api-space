import requests
import os
import xlsxwriter
from datetime import datetime
TOKEN = os.getenv('TOKEN')
URL = 'https://vortexz.jetbrains.space/api/http/projects/key:EVENLINEZ/planning/issues?sorting=CREATED&' \
      'descending=true&$fields=data(status(name,id),title,sprints(board(board(id,name)),description),' \
      'creationTime,description)'


def expor_issue():
    headers = {'Authorization': 'Bearer ' + TOKEN, 'Accept': 'application/json'}
    datas = requests.get(URL, headers=headers).json()["data"]
    for x in datas:
        for y in x["sprints"]:
            if y["board"]["board"]["name"] == '2021':
                dates = [
                    {"titulo": x["title"]},
                    {"description": x["description"]},
                    {"board": y["board"]["board"]["name"]},
                    {"created": x["creationTime"]["iso"]},
                    {"Status": x["status"]["name"]}
                ]
                return dates


def export_excel():
    book = xlsxwriter.Workbook('Issues.xlsx')
    sheet = book.add_worksheet("Issues BOARD")
    sheet.merge_range('A1:F1', "SPRINT", book.add_format({'bold': True, 'font_size': 14, 'align': 'center',
                                                          'valign': 'vcenter', 'font_name': 'Arial',
                                                          'color': '#171717'}))
    bold = book.add_format({'bold': 1})
    date_format = book.add_format({'num_format': 'mmmm d yyyy'})
    sheet.set_column(1, 1, 15)
    sheet.write('A2', 'Title', bold)
    sheet.write('B2', 'Description', bold)
    sheet.write('C2', 'Created', bold)
    sheet.write('D2', 'Board', bold)
    sheet.write('E2', 'Status', bold)

    expenses = (
        ['Rent', '2013-01-13', 1000],
        ['Gas', '2013-01-14', 100],
        ['Food', '2013-01-16', 300],
        ['Gym', '2013-01-20', 50],
    )

    row = 2
    col = 0
    for item, date_str, cost in (expenses):
        date = datetime.strptime(date_str, "%Y-%m-%d")

        sheet.write_string(row, col, item)
        sheet.write_string(row, col, item)
        sheet.write_datetime(row, col + 1, date, date_format)
        sheet.write_number(row, col + 2, cost)
        sheet.write_string(row, col, item)
        row += 1

    book.close()


expor_issue()
# export_excel()
