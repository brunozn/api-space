import requests
import os
TOKEN = os.getenv('TOKEN')
URL = 'https://vortexz.jetbrains.space/api/http/projects/key:EVENLINEZ/planning/issues?sorting=CREATED&' \
      'descending=true&$fields=data(status(name,id),title,sprints(board(board(id,name)),description),' \
      'creationTime,description)'


def expor_issue():
    url = URL
    headers = {'Authorization': 'Bearer ' + TOKEN, 'Accept': 'application/json'}
    datas = requests.get(url, headers=headers).json()["data"]
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
                print(dates)


expor_issue()
