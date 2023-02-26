from __future__ import print_function

from apiclient import discovery
from httplib2 import Http
from oauth2client import client, file, tools

SCOPES = "https://www.googleapis.com/auth/forms.body"
DISCOVERY_DOC = "https://forms.googleapis.com/$discovery/rest?version=v1"

store = file.Storage('token.json')
creds = None
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('/Users/suzukiharumasaru/apps/WG-RPA/seacrets/credentials.json', SCOPES)
    creds = tools.run_flow(flow, store)

form_service = discovery.build('forms', 'v1', http=creds.authorize(
    Http()), discoveryServiceUrl=DISCOVERY_DOC, static_discovery=False)

# Request body for creating a form
NEW_FORM = {
    "info": {
        "title": "Quickstart form",
    }
}

# Request body to add a multiple-choice question
NEW_QUESTION = {
    "requests": [{
        "createItem": {
            "item": {
                "title": "質問①",
                "questionItem": {
                    "question": {
                        "required": True,
                        "choiceQuestion": {
                            "type": "RADIO",
                            "options": [
                                {"value": "回答1"},
                                {"value": "回答2"},
                                {"value": "回答3"},
                                {"value": "回答4"}
                            ],
                            "shuffle": True
                        }
                    }
                },
            },
            "location": {
                "index": 0
            }
        }
    }]
}

# Creates the initial form
result = form_service.forms().create(body=NEW_FORM).execute()

# Adds the question to the form
question_setting = form_service.forms().batchUpdate(formId=result["formId"], body=NEW_QUESTION).execute()

# Prints the result to show the question has been added
get_result = form_service.forms().get(formId=result["formId"]).execute()
print(get_result)