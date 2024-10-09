import base64
from email.message import EmailMessage
import os.path
import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

def gmail_create_draft():
  """Create and insert a draft email.
   Print the returned draft's message and id.
   Returns: Draft object, including draft id and message meta data.

  Load pre-authorized user credentials from the environment.
  TODO(developer) - See https://developers.google.com/identity
  for guides on implementing OAuth2 for the application.
  """
  creds = None
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
    # create gmail api client
    service = build("gmail", "v1", credentials=creds)

    message = EmailMessage()

    message.set_content("This is automated draft mail")

    message["To"] = "gduser1@workspacesamples.dev"
    message["From"] = "filip.mihai.muntean@gmail.com"
    message["Subject"] = "Automated draft"

    # encoded message
    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    create_message = {"message": {"raw": encoded_message}}
    # pylint: disable=E1101
    draft = (
        service.users()
        .drafts()
        .create(userId="me", body=create_message)
        .execute()
    )

    print(f'Draft id: {draft["id"]}\nDraft message: {draft["message"]}')

  except HttpError as error:
    print(f"An error occurred: {error}")
    draft = None

  return draft


if __name__ == "__main__":
  gmail_create_draft()