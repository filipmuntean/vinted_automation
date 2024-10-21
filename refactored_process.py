import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64
from bs4 import BeautifulSoup
from datetime import datetime

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.readonly']
SPREADSHEET_ID = '1oBMj-n4iRuDqGmbKpBeB3P_qL4ucAIhr-phU40gsRRA'
SHEET_RANGE = 'Sheet1!C:D'
CURRENT_DATE = datetime.now().strftime("%Y-%m-%d")


### Class 1: GmailManager

class GmailManager:
    def __init__(self):
        self.creds = None
        self.service = self.authenticate_gmail_api()

    def authenticate_gmail_api(self):
        if os.path.exists('token.json'):
            self.creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())
        return build("gmail", "v1", credentials=self.creds)

    def fetch_emails(self, query='from:@vinted.nl'):
        try:
            results = self.service.users().messages().list(userId="me", q=query).execute()
            return results.get("messages", [])
        except HttpError as error:
            print(f"An error occurred: {error}")
            return []

    def get_message_subject(self, msg):
        for header in msg['payload']['headers']:
            if header['name'] == 'Subject':
                return header['value']
        return "No Subject"

    def get_message_body(self, msg):
        if 'parts' in msg['payload']:
            for part in msg['payload']['parts']:
                if part['mimeType'] == 'text/plain':
                    return base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
                elif part['mimeType'] == 'text/html':
                    html_body = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
                    soup = BeautifulSoup(html_body, 'html.parser')
                    return soup.get_text()
        else:
            return base64.urlsafe_b64decode(msg['payload']['body']['data']).decode('utf-8')
        return None

    def extract_item_and_price(self, html_body):
        soup = BeautifulSoup(html_body, 'html.parser')
        item_name = ""
        sale_info = soup.find_all('p')
        for p_tag in sale_info:
            if "Your sale of" in p_tag.get_text():
                item_name_span = p_tag.find('span')
                if item_name_span:
                    item_name = item_name_span.get_text(strip=True).split("was completed")[0].strip()
                break
        price_info = soup.find('td', string='Item price:')
        price = price_info.find_next('td').get_text(strip=True) if price_info else ""
        return item_name, price


### Class 2: SheetsManager

class GoogleSheets:
    def __init__(self):
        self.creds = None
        self.service = self.authenticate_sheets_api()

    def authenticate_sheets_api(self):
        if os.path.exists('token.json'):
            self.creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())
        return build('sheets', 'v4', credentials=self.creds)

    def check_if_total_exists(self, sheet_title, current_row_count):
        result = self.service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_title}!C{current_row_count-1}:C{current_row_count}"
        ).execute()
        values = result.get('values', [])
        return any("Total" in row[0] for row in values if row)

    def get_next_empty_row(self, sheet_title):
        result = self.service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, 
            range=f"{sheet_title}!A:E"
        ).execute()
        values = result.get('values', [])
        return len(values) + 1

    def append_new_items_in_decreasing_order(self, sheet_title, new_items):
        result = self.service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"{sheet_title}!C:C").execute()
        existing_items = [row[0] for row in result.get('values', []) if row]
        missing_items = [item for item in new_items if item['name'] not in existing_items]
        if not missing_items:
            print("No new items to append.")
            return
        missing_items.reverse()
        next_row = self.get_next_empty_row(sheet_title)
        for item in missing_items:
            formula = f"=D{next_row} - B{next_row} / 5"
            values = [[None, None, item['name'], item['price'], formula]]
            body = {'values': values}
            self.service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{sheet_title}!A{next_row}:E{next_row}",
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            print(f"Appended: {item['name']} - {item['price']} €")
            next_row += 1

    def handle_totals_and_new_sheet(self, sheet_title, new_items):
        result = self.service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range="Sheet1!A:E").execute()
        
        current_row_count = len(result.get('values', []))
        if current_row_count >= 50:
            if not self.check_if_total_exists(sheet_title, current_row_count):
                total_row = [[None, None, "Total", f"=SUM(D1:D{current_row_count})", f"=SUM(E1:E{current_row_count})"]]
                self.service.spreadsheets().values().append(
                    spreadsheetId=SPREADSHEET_ID,
                    range=f"Sheet1!A{current_row_count + 1}:E{current_row_count + 1}",
                    valueInputOption='USER_ENTERED',
                    body={'values': total_row}
                ).execute()
                print("Total row appended. Now creating a new sheet.")

            new_sheet_title = f"Sheet_{CURRENT_DATE}"
            new_sheet_id = self.get_sheet_by_title(new_sheet_title)
            if new_sheet_id is None:
                create_sheet_body = {'requests': [{'addSheet': {'properties': {'title': new_sheet_title, 'gridProperties': {'rowCount': 100, 'columnCount': 10}}}}]}
                self.service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=create_sheet_body).execute()
                print(f"New sheet '{new_sheet_title}' created.")
            self.append_new_items_in_decreasing_order(new_sheet_title, new_items)
        else:
            self.append_new_items_in_decreasing_order(sheet_title, new_items)


### Main Process:

def main():
    gmail_manager = GmailManager()
    sheets_manager = GoogleSheets()

    emails = gmail_manager.fetch_emails()
    if not emails:
        print("No emails found.")
        return

    for email in emails:
        msg = gmail_manager.service.users().messages().get(userId='me', id=email['id']).execute()
        subject = gmail_manager.get_message_subject(msg)
        if subject == "This order is completed":
            body = gmail_manager.get_message_body(msg)
            item, price = gmail_manager.extract_item_and_price(body)
            sheets_manager.handle_totals_and_new_sheet("Sheet1", [{'name': item, 'price': float(price[:-2])}])

if __name__ == "__main__":
    main()
