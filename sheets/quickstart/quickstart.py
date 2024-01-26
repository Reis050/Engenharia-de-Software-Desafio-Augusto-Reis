import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of spreadsheet.
SPREADSHEET_ID = "1Lp2k-ZstPwuHw_iCK7E_DSxz7mGL03fc6AkI7QKnqyg"
FAULTS_RANGE = "engenharia_de_software!C4:C27"
NOTES_RANGE = "engenharia_de_software!D4:F27"
lessons_total=60

def main():
    
# google api code 
  creds = None 
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    
            
    # values to update according to each situation
    aprove=[
        ["Aprovado", 0]
    ]
    
   
    repByNotes=[
        ["Reprovado por Nota", 0]
    ]
    
    # working with the spreadsheet only with the note cells
    result = (
        sheet.values()
        .get(spreadsheetId=SPREADSHEET_ID, range=NOTES_RANGE)
        .execute()
    )
    n_values = result.get("values", [])
    
    # adding up the notes  
    sums = [sum(map(int, line)) for line in n_values]
    
    # scrolling and updating according to each situation
    for i in range(len(sums)):
        m=sums[i]/30
        # final exam
        if (m>=5 and m<7 ):
            naf=10-m
            finalEx=[["Exame Final", round(naf, 2)]]
            result = (
            sheet.values()
            .update(spreadsheetId=SPREADSHEET_ID, range=f"engenharia_de_software!G{i+4}", valueInputOption="USER_ENTERED", body={"values":finalEx}, )
            .execute()
            )
        # approved
        elif (m>=7 ):
            result = (
            sheet.values()
            .update(spreadsheetId=SPREADSHEET_ID, range=f"engenharia_de_software!G{i+4}", valueInputOption="USER_ENTERED", body={"values":aprove}, )
            .execute()
            )
        # failed   
        elif (m<5 ):
            result = (
            sheet.values()
            .update(spreadsheetId=SPREADSHEET_ID, range=f"engenharia_de_software!G{i+4}", valueInputOption="USER_ENTERED", body={"values":repByNotes}, )
            .execute()
            )   
    
    # working with the spreadsheet only with the faults cells,
    result = (
        sheet.values()
        .get(spreadsheetId=SPREADSHEET_ID, range=FAULTS_RANGE)
        .execute()
    )
    f_values = result.get("values", [])
    repByFault=[
        ["Reprovado por Falta", 0]
    ]
    # updating and overwriting the situation of those who failed due to absence
    for i, value in enumerate(f_values, 4):
        if int(value[0]) > (0.25 * lessons_total):
            result = (
        sheet.values()
        .update(spreadsheetId=SPREADSHEET_ID, range=f"engenharia_de_software!G{i}", valueInputOption="USER_ENTERED", body={"values":repByFault}, )
        .execute()
    )
            
                
  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()