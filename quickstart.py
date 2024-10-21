import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64
from bs4 import BeautifulSoup
from datetime import datetime

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.readonly']

SPREADSHEET_ID = '1oBMj-n4iRuDqGmbKpBeB3P_qL4ucAIhr-phU40gsRRA'

SHEET_RANGE = 'Sheet1!C:D'
CURRENT_DATE = datetime.now().strftime("%Y-%m-%d")

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


def check_if_total_exists(service, spreadsheet_id, sheet_title, current_row_count):
    """Check if the 'Total' row already exists in the given sheet."""
    # Check the last few rows for the word 'Total' in column C
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, 
        range=f"{sheet_title}!C{current_row_count-1}:C{current_row_count}"  # Check last two rows in column C
    ).execute()

    values = result.get('values', [])
    
    for row in values:
        if row and "Total" in row[0]:  # If 'Total' is found in column C, return True
            return True
    return False

def get_sheet_by_title(service, spreadsheet_id, sheet_title):
    """Check if a sheet with the given title already exists and return its sheetId."""
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet.get('sheets', [])
    
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_title:
            return sheet['properties']['sheetId']
    return None

def get_next_empty_row(service, spreadsheet_id, sheet_title):
    """Find the next empty row in the given sheet."""
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, 
        range=f"{sheet_title}!A:E"  # Check range A to E to find the last row with data
    ).execute()

    values = result.get('values', [])
    return len(values) + 1  # The next empty row will be after the last row with data


def get_last_existing_items(service, spreadsheet_id, sheet_title, num_rows=10):
    """Retrieve the last `num_rows` from the sheet to check for already existing items."""
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, 
        range=f"{sheet_title}!C:C"  # Check column C (which contains item names)
    ).execute()

    values = result.get('values', [])
    # Get the last `num_rows` items from column C
    return [row[0] for row in values[-num_rows:] if row]  # Return a list of item names


def append_new_items_in_decreasing_order(service, spreadsheet_id, sheet_title, new_items):
    """Append only new items in decreasing order to the sheet."""
    # Get the last existing items in the sheet (e.g., last 10 rows)
    last_existing_items = get_last_existing_items(service, spreadsheet_id, sheet_title)

    # Find the items that are missing from the sheet
    missing_items = [item for item in new_items if item['name'] not in last_existing_items]
    # If there are no new items, we skip appending
    if not missing_items:
        print("No new items to append.")
        return

    # Add the missing items in decreasing order (reverse the list)
    missing_items.reverse()
    print(missing_items)
    # Append the missing items one by one
    next_row = get_next_empty_row(service, spreadsheet_id, sheet_title)  # Get the next available row
    for item in missing_items:
        formula = f"=D{next_row} - B{next_row} / 5"
        values = [[None, None, item['name'], item['price'], formula]]  # Item name in C, price in D, formula in E
        body = {
            'values': values
        }

        # Append the new data in the next available row
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_title}!A{next_row}:E{next_row}",
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()

        print(f"Appended: {item['name']} - {item['price']} €")
        next_row += 1  # Move to the next row for the next item

def append_to_google_sheets(item, price):
    """Insert a new row with item, price, and formula in decreasing order."""
    service = authenticate_sheets_api()

    # Define the current sheet title
    sheet_title = "Sheet1"

    # Check if there are existing items already in the sheet
    new_items = [{'name': item, 'price': price}]
    
    # Retrieve the last existing items in the sheet (e.g., last 10 rows)
    last_existing_items = get_last_existing_items(service, SPREADSHEET_ID, sheet_title)

    # Find the items that are missing from the sheet
    missing_items = [item for item in new_items if item['name'] not in last_existing_items]

    # If there are no new items, skip appending
    if not missing_items:
        print("No new items to append.")
        return

    # If the sheet has reached 50 rows, calculate totals and move to a new sheet
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Sheet1!A:E").execute()
    existing_rows = result.get('values', [])
    current_row_count = len(existing_rows)

    if current_row_count >= 50:
        if not check_if_total_exists(service, SPREADSHEET_ID, sheet_title, current_row_count):
            # Compute totals only after reaching 50 rows
            total_column_d = f"=SUM(D1:D{current_row_count})"
            total_column_e = f"=SUM(E1:E{current_row_count})"

            # Append the totals to the next available row in the current sheet
            total_row = [[None, None, "Total", total_column_d, total_column_e]]
            sheet.values().append(
                spreadsheetId=SPREADSHEET_ID,
                range=f"Sheet1!A{current_row_count + 1}:E{current_row_count + 1}",
                valueInputOption='USER_ENTERED',
                body={'values': total_row}
            ).execute()

            print("Total row appended. Now creating a new sheet.")
        else:
            print("Already have the total row.")

        # Generate a new sheet name using the current date
        current_date = datetime.now().strftime("%Y-%m-%d")
        new_sheet_title = f"Sheet_{current_date}"

        # Check if the new sheet already exists
        new_sheet_id = get_sheet_by_title(service, SPREADSHEET_ID, new_sheet_title)

        if new_sheet_id is None:
            # Create the new sheet if it doesn't exist
            create_sheet_body = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': new_sheet_title,
                            'gridProperties': {'rowCount': 100, 'columnCount': 10}
                        }
                    }
                }]
            }

            # Execute the creation of a new sheet
            response = sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=create_sheet_body).execute()
            new_sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']
            print(f"New sheet '{new_sheet_title}' created.")
        else:
            print(f"Sheet '{new_sheet_title}' already exists. Moving to it.")

        # Now, append the missing items to the new sheet in decreasing order
        append_new_items_in_decreasing_order(service, SPREADSHEET_ID, new_sheet_title, missing_items)

    else:
        # If we're still under the 50-row limit, append the new items to the existing sheet in decreasing order
        append_new_items_in_decreasing_order(service, SPREADSHEET_ID, sheet_title, missing_items)

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
    sale_info = soup.find_all('p')

    for p_tag in sale_info:
      if "Your sale of" in p_tag.get_text():  
          item_name_span = p_tag.find('span')
          if item_name_span:
              item_name = item_name_span.get_text(strip=True).split("was completed")[0].strip()  
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
