import requests
import os
import xlsxwriter

# from datetime import datetime
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
                value = [(
                    x["title"], x["description"],
                    x["creationTime"]["iso"], x["status"]["name"]
                )]

                dates = [
                    {'Title': x["title"], 'Description': x["description"], 'Created':  x["creationTime"]["iso"], 'Status': x["status"]["name"]},
                ]

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

                example = [{'Title': 'MG1', 'Description': 'SP1', 'Created': 'RJ2', 'Status': 'CE4'},
                           {'Title': 'MG2', 'Description': 'SP2', 'Created': 'RJ3', 'Status': 'CE3'},
                           {'Title': 'MG3', 'Description': 'SP3', 'Created': 'RJ4', 'Status': 'CE2'},
                           {'Title': 'MG4', 'Description': 'SP4', 'Created': 'RJ5', 'Status': 'CE1'},
                           ]
                for idy, data in enumerate(dates):
                    print(idy, data)
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


expor_issue()
# export_excel(r)
