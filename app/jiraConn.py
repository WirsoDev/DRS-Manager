import os
import dotenv
from jira import JIRA
from consts import requests_types

dotenv.load_dotenv()

JIRA_URL = os.environ.get('JIRA_URL')
JIRA_USER = os.environ.get('JIRA_USER')
JIRA_TOKEN = os.environ.get('JIRA_TOKEN')


class JiraConn:
    def __init__(self):
        self.url = JIRA_URL
        self.user = JIRA_USER
        self.token = JIRA_TOKEN
        self.jira = JIRA(server=self.url, basic_auth=(self.user, self.token))
        self.proj = 'DM'

    def treat_drs(sefl, data:dict):
        drs_number = data['drsNumber']
        drs_version = data['drsVersion']
        client = data['client']
        requests = data['tiposPedido']

        print(requests)
        
        #chech if requests is a list by (requests.split(',))
        if isinstance(requests, str):
            requests = requests.split(',')

        #check if request is MRK
        r_request = []
        for r in requests:
            r = r.strip()
            if r in requests_types['MRK']:
                r_request.append(r)
        
        if len(r_request) == 0:
            return False

        #if len requests > 1 create subtasks
        if len(r_request) > 1:
            sub_tasks = []
            for request in r_request:
                sub_tasks.append({"summary": request, "description": request})
        else:
            sub_tasks = None
            r_request = r_request[0].replace(',', '').strip()
        
        file = fr"\\synnas\CREATIVO\DRS\DRS_PDF\DRS_ {drs_number}_ED_ {drs_version}.pdf"

        data = {
            'drs_number':drs_number,
            'drs_version':drs_version,
            'client':client,
            'requests':r_request,
            'sub_tasks':sub_tasks,
            'file':[file]
            }
        print(data)
        return data


    def create_issue(
        self, 
        summary, 
        description, 
        issuetype: dict[str, str], 
        sub_tasks: list[dict[str, str]], 
        attachments: list[str] = None,
    ):
        print(sub_tasks)
        
        resp = ''
        '''
        sub_tasks = [
            {"summary": "Sub-task 1", "description": "Details of sub-task 1"},
            {"summary": "Sub-task 2", "description": "Details of sub-task 2"},
            {"summary": "Sub-task 3", "description": "Details of sub-task 3"},
        ]
        
        attachments = [
            "path/to/file1.txt",
            "path/to/file2.jpg"
        ]
        '''
        
        try:
            print('adding issue...')
            # Create the parent issue
            parent_issue = self.jira.create_issue(
                project=self.proj,
                summary=summary,
                description=description,
                issuetype=issuetype,
            )

            resp += parent_issue.key

            # Attach files to the parent issue (if provided)
            if attachments:
                for file_path in attachments:
                    try:
                        with open(file_path, "rb") as file:
                            self.jira.add_attachment(issue=parent_issue, attachment=file)
                    except Exception as attach_error:
                        print(f"Error attaching file '{file_path}': {attach_error}")

            # Create sub-tasks
            if sub_tasks:
                print('adding subtasks...')
                for task in sub_tasks:
                    sub_issue = self.jira.create_issue(
                        project=self.proj,
                        summary=task["summary"],
                        description=task["description"],
                        issuetype={"name": "Sub-task"},
                        parent={"key": parent_issue.key}
                    )
                    print(f"Created Sub-Task: {sub_issue.key} \n")

            return resp
        except Exception as e:
            return e


    def run_jira(self, data):
        print('     -> Run jira...')
        drs_to_add = self.treat_drs(data)
        if not drs_to_add:
            print('     -> No tasks o add!')
            return False
        
        summary = f'DRS {drs_to_add['drs_number']}_{drs_to_add['drs_version']} - {drs_to_add['client']}'
        description = f'{drs_to_add['requests']}'
        sub_tasks = drs_to_add['sub_tasks']
        file = drs_to_add['file']

        resp = self.create_issue(
            summary=summary,
            description=description,
            issuetype={'name': 'Task'},
            sub_tasks=sub_tasks,
            attachments=file
        )
        print(f'     -> jira - {resp}')

        return resp