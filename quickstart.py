import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64
from bs4 import BeautifulSoup

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

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
            
        print(f"Item: {item}\nPrice: {price}\n{'-'*50}\n")
        # print(f"{body}\n\n\n")

  except HttpError as error:
    # TODO(developer) - Handle errors from gmail API.
    print(f"An error occurred: {error}")


if __name__ == "__main__":
  main()
