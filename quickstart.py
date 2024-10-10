import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64
from bs4 import BeautifulSoup

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.readonly']

SPREADSHEET_ID = '1oBMj-n4iRuDqGmbKpBeB3P_qL4ucAIhr-phU40gsRRA'

SHEET_RANGE = 'Sheet1!C:D'

def authenticate_sheets_api():
    """Authenticate and return the Google Sheets API service."""
    creds = None
    # The token.json stores the user's access and refresh tokens.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for future runs
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # Build the Sheets API service
    service = build('sheets', 'v4', credentials=creds)
    return service

def append_to_google_sheets(item, price):
    """Insert a new row with item, price, and formula at the top of the Google Sheets."""
    service = authenticate_sheets_api()

    # Retrieve current data to know where to insert the new row
    sheet = service.spreadsheets()

    # Insert a blank row at the top (shift everything down)
    request_body = {
        'requests': [{
            'insertRange': {
                'range': {
                    'sheetId': 0,  # Assuming first sheet (sheetId=0), adjust if needed
                    'startRowIndex': 0,  # Start at the top
                    'endRowIndex': 1  # Insert only one row
                },
                'shiftDimension': 'ROWS'
            }
        }]
    }

    # Send the request to insert the row
    sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=request_body).execute()

    # Prepare the data to insert in the first row
    formula = f"=D1 - B1/ 5"
    values = [[None, None, item, price, formula]]  # None for A, B columns, item in C, price in D, formula in E
    body = {
        'values': values
    }

    # Append item, price, and formula in the first row (C1:D1:E1)
    result = sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range="Sheet1!A1:E1",  # Insert at the top row (first row)
        valueInputOption='USER_ENTERED',  # Ensure formula is entered correctly
        body=body
    ).execute()

    print(f'{result.get("updatedCells")} cells updated.')

# def append_to_google_sheets(item, price):
#     """Append a new row with item, price, and formula to the Google Sheets."""
#     service = authenticate_sheets_api()

#     # Retrieve current data to know where to append new rows
#     sheet = service.spreadsheets()

#     # Get the existing rows in the sheet to calculate the next available row
#     result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Sheet1!A:E").execute()
#     existing_rows = result.get('values', [])
#     next_row = len(existing_rows) + 1  # Calculate the next available row

#     # Create the formula for column E
#     formula = f"=D{next_row} - B{next_row}/ 5"

#     # Prepare the data to append, item in column C, price in column D, and formula in column E
#     values = [[None, None, item, price, formula]]  # None for A, B columns, item in C, price in D, formula in E
#     body = {
#         'values': values
#     }

#     # Append item, price, and formula to Google Sheets
#     result = sheet.values().append(
#         spreadsheetId=SPREADSHEET_ID,
#         range=f"Sheet1!A{next_row}:E{next_row}",
#         valueInputOption='USER_ENTERED',  # This ensures the formula is entered correctly
#         body=body
#     ).execute()

#     print(f'{result.get("updates").get("updatedCells")} cells appended.')


def get_message_subject(msg):
    """Extracts the subject (title) of the email from the headers."""
    for header in msg['payload']['headers']:
        if header['name'] == 'Subject':
            return header['value']
    return "No Subject"

def extract_item_and_price(html_body):
    """Extracts the item name and price from the HTML email body."""
    soup = BeautifulSoup(html_body, 'html.parser')

    # Extract item name (searching for the first occurrence of "Your sale of")
    item_name = ""
    # sale_info = soup.find('p', string='Your sale of&nbsp;')
    sale_info = soup.find_all('p')
    # print(sale_info, "UUUU")

    for p_tag in sale_info:
      if "Your sale of" in p_tag.get_text():  
          item_name_span = p_tag.find('span')
          if item_name_span:
              item_name = item_name_span.get_text(strip=True).split("was completed")[0].strip()  # Extract the item name
              break  # We found the item, exit the loop
    
    # Extract price (searching for "Item price")
    price = ""
    price_info = soup.find('td', string='Item price:')
    if price_info:
        price = price_info.find_next('td').get_text(strip=True)
    
    return item_name, price

def get_message_body(msg):
    """Extracts the body of the message, handling different types of content."""
    if 'parts' in msg['payload']:
        # If the message is multipart, extract each part
        for part in msg['payload']['parts']:
            if part['mimeType'] == 'text/plain':
                print("HERE")
                return base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
            elif part['mimeType'] == 'text/html':
                html_body = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
                soup = BeautifulSoup(html_body, 'html.parser')
                return soup.get_text()
    else:
        # If the message is not multipart, extract from the single part
        return base64.urlsafe_b64decode(msg['payload']['body']['data']).decode('utf-8')
    return None


def main():
  """Shows basic usage of the Gmail API.
  Lists the user's Gmail labels.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    # Call the Gmail API
    service = build("gmail", "v1", credentials=creds)

    query = 'from:@vinted.nl'

    results = service.users().messages().list(userId="me", q = query).execute()
    labels = results.get("messages", [])

    if not labels:
      print("No labels found.")
      return
    
    print("Labels from Vinted:")
    for message in labels:
      lbl = service.users().messages().get(userId='me', id=message['id']).execute()

      subject = get_message_subject(lbl)
      # print(f"Message snippet: {lbl['snippet']}")
      body = get_message_body(lbl)

      if subject == "This order is completed":
        item, price = extract_item_and_price(body)

        if price.startswith('\''):
          price = price[1:]
        append_to_google_sheets(item, float(price[:-2]))
        print(f"Added to Google Sheets: {item} - {price}")
        
  except HttpError as error:
    # TODO(developer) - Handle errors from gmail API.
    print(f"An error occurred: {error}")


if __name__ == "__main__":
  main()
